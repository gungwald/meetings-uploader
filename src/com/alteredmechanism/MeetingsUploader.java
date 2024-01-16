package com.alteredmechanism;

import com.microsoft.outlook.*;
import com4j.ErrorInfo;

import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpRequest.BodyPublisher;
import java.net.http.HttpRequest.BodyPublishers;
import java.net.http.HttpResponse;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;

import static com.microsoft.outlook.OlDefaultFolders.olFolderCalendar;
import static java.util.Calendar.*;

public class MeetingsUploader {

    public static Logger logger = Logger.getLogger(MeetingsUploader.class.getName());
    public static URI SEND_TO_ADDRESS;

    public static void main(String[] args) {
        try {
            if (args.length > 0) {
                SEND_TO_ADDRESS = new URI(args[0]);
                MeetingsUploader mu = new MeetingsUploader();
                mu.upload(SEND_TO_ADDRESS);
            } else {
                System.err.println("You stupid");
            }
        } catch (Throwable t) {
            logger.log(Level.SEVERE, "Failed to upload meetings", t);
        }
    }

    protected void upload(URI uploadAddress) throws IOException, InterruptedException {
        _Application lookout = ClassFactory.createApplication();
        _NameSpace ns = lookout.getNamespace("MAPI");
        MAPIFolder calendar = ns.getDefaultFolder(olFolderCalendar);
        _Items allMeetings = calendar.getItems();
        ArrayList<_AppointmentItem> meetings = new ArrayList<>();
        Date now = new Date();

        _AppointmentItem meeting = allMeetings.getFirst().queryInterface(_AppointmentItem.class);

        while (meeting != null) {
            if (! meeting.getAllDayEvent() && ! isOptional(meeting)) {
                if (meeting.getIsRecurring()) {
                    meetings.addAll(resolveRecurrences(now, meeting, 14));
                } else {
                    if (meeting.getStart().after(now)) {
                        meetings.add(meeting);
                    }
                }
            }
            meeting = allMeetings.getNext().queryInterface(_AppointmentItem.class);
        }
        String json = convertToJson(meetings);
        sendViaHttpPut(json, uploadAddress);
    }

    protected void sendViaHttpPut(String json, URI sendToAddress) throws IOException, InterruptedException {
        logger.info("Sending meetings:" + json);
        HttpClient http = HttpClient.newBuilder().build();
        BodyPublisher reqBody = BodyPublishers.ofString(json);
        HttpRequest req = HttpRequest.newBuilder().PUT(reqBody).uri(sendToAddress).build();
        HttpResponse.BodyHandler<String> respBodyHandler = HttpResponse.BodyHandlers.ofString();
        HttpResponse<String> resp = http.send(req, respBodyHandler);
        logger.info("Received:" + resp.body());
    }

    private String convertToJson(ArrayList<_AppointmentItem> meetings) {
        SimpleDateFormat dateTimeFmt = new SimpleDateFormat("yyyy-MM-dd HH:mm");
        StringBuilder json = new StringBuilder("{\"meetings\":[\r\n");
        for (_AppointmentItem meeting : meetings) {
            String startDateTime = dateTimeFmt.format(meeting.getStart());
            String subject;
            if (meeting.getConversationTopic() != null) {
                subject = encodeJsonValue(meeting.getConversationTopic());
            } else {
                subject = "";
            }
            json.append("{\"dateTime\":\"");
            json.append(startDateTime);
            json.append("\",\"subject\":\"");
            json.append(subject);
            json.append("\"},\n");
        }
        return json.toString();
    }

    private String encodeJsonValue(String conversationTopic) {
        StringBuilder encoded = new StringBuilder();
        for (char c : conversationTopic.toCharArray()) {
            switch (c) {
                case '"':
                    encoded.append("\\\"");
                case '\b':
                    encoded.append("\\b");
                case '\r':
                    encoded.append("\\r");
                case '\n':
                    encoded.append("\\n");
                case '\t':
                    encoded.append("\\t");
                case '\\':
                    encoded.append("\\");
                default:
                    encoded.append(c);
            }
        }
        return encoded.toString();
    }

    protected boolean isOptional(_AppointmentItem meeting) {
        return meeting.getConversationTopic().toLowerCase().contains("(optional)");
    }

    protected ArrayList<_AppointmentItem> resolveRecurrences(Date beginningWithDate, _AppointmentItem meeting, int dayCount) {
        ArrayList<_AppointmentItem> recurrences = new ArrayList<>();
        // Take date from targetDateTime and time from meeting.Start
        Calendar toCheck = joinDateAndTime(beginningWithDate, meeting.getStart());
        for (int i = 0; i < dayCount; i++) {
            _AppointmentItem recurrence = resolveRecurrence(toCheck.getTime(), meeting.getRecurrencePattern());
            if (recurrence != null) {
                recurrences.add(recurrence);
            }
            toCheck.add(DAY_OF_MONTH, 1);
        }
        return recurrences;
    }

    private Calendar joinDateAndTime(Date useForDate, Date useForTime) {
        // Convert Date to Calendar
        Calendar forDate = Calendar.getInstance();
        forDate.setTime(useForDate);
        // Convert Date to Calendar
        Calendar forTime = Calendar.getInstance();
        forTime.setTime(useForTime);

        // Build the joined date and time Calendar
        Calendar combined = Calendar.getInstance();
        combined.clear();
        // Date values
        combined.set(YEAR, forDate.get(YEAR));
        combined.set(MONTH, forDate.get(MONTH));
        combined.set(DAY_OF_MONTH, forDate.get(DAY_OF_MONTH));
        // Time values
        combined.set(HOUR_OF_DAY, forTime.get(HOUR_OF_DAY));
        combined.set(MINUTE, forTime.get(MINUTE));
        return combined;
    }

    protected _AppointmentItem resolveRecurrence(Date targetDateTime, RecurrencePattern recPattern) {
        _AppointmentItem occurrence = null;

        try {
            occurrence = recPattern.getOccurrence(targetDateTime);
        } catch (com4j.ComException e) {

            ErrorInfo err = e.getErrorInfo();
            if (err != null) {
                logger.warning(err.toString());
            }
            logger.log(Level.WARNING, "HRESULT={}", e.getHRESULT());
            logger.log(Level.WARNING, e.getDetailMessage());
            logger.warning(e.toString());
            e.printStackTrace();
        }
        return occurrence;
    }
}


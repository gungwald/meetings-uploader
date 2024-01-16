package com.alteredmechanism;

import com.microsoft.outlook.*;
import com4j.Com4jObject;
import com4j.ComException;

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

        _AppointmentItem meeting = getFirst(allMeetings);

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
            meeting = getNext(allMeetings);
        }
        String json = convertToJson(meetings);
        sendViaHttpPut(json, uploadAddress);
    }

    protected _AppointmentItem getFirst(_Items items) {
        _AppointmentItem meeting;
        Com4jObject item = items.getFirst();
        if (item != null) {
            meeting = item.queryInterface(_AppointmentItem.class);
        } else {
            meeting = null;
        }
        return meeting;
    }

    protected _AppointmentItem getNext(_Items items) {
        _AppointmentItem meeting;
        Com4jObject item = items.getNext();
        if (item != null) {
            meeting = item.queryInterface(_AppointmentItem.class);
        } else {
            meeting = null;
        }
        return meeting;
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
        return meeting.getConversationTopic() != null && meeting.getConversationTopic().toLowerCase().contains("(optional)");
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
        final int NO_RECURRENCE_ERR_NUM = -2147467259;

        try {
            occurrence = recPattern.getOccurrence(targetDateTime);
        } catch (ComException e) {
            if (e.getHRESULT() != NO_RECURRENCE_ERR_NUM && e.getHRESULT() != 0) {
                logger.log(Level.SEVERE, String.format("Failed to get recurrence: VBA/VBS Err.Number a.k.a. C++ HRESULT = %d: %s", e.getHRESULT(), targetDateTime.toString()), e);
            }
        }
        return occurrence;
    }
}


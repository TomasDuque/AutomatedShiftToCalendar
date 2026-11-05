# Automated Shift Scheduler

A Google Apps Script that turns my restaurant's weekly Excel schedule into Google Calendar events — automatically, every week.

Every Thursday or Friday my manager emails the next week's schedule as an `.xlsx` attachment. I used to copy each shift into my calendar by hand and then reopen the spreadsheet to see who I was working with. This script does that for me. It finds the email, reads the spreadsheet, and creates a calendar event for each of my shifts with everything I need to know about it.


## What it does

- **Finds the schedule email** in Gmail (read-only) by sender, subject, and attachment
- **Parses the Excel grid**: a week of AM/PM columns per day, split into sections for managers, bartenders, servers, hosts, and barbacks
- **Creates a calendar event for each of my shifts** with the restaurant's address as the location and an estimated end time
- **Writes a description** listing everyone working that day, grouped by role and sorted by start time, plus private events and parties for that day
- **Flags key-hourly shifts** (when I'm the person in charge) in the event title, and shows a timeline of who's in charge through the day
- **Handles revised schedules**: if a corrected schedule is sent, that week's events are replaced instead of duplicated
- **Runs on a timer** (every 8 hours), so new schedules show up without me doing anything

## Example output

What the description looks like for one shift, generated from the sample schedule in this repo (`node test/run-sample.js`):

```
=== Work ⭐ KEY HOURLY | Sat, Nov 7, 4:00 PM -> Sun, Nov 8, 12:00 AM
⭐ YOU ARE THE KEY HOURLY from 6pm (manager on duty before that)
Your shift: 4pm – 12am (closing)

WHO'S IN CHARGE TODAY
• 10am–12pm: Luke (key hourly)
• 12pm–6pm: Han (manager)
• 6pm–12am: Me (key hourly)

MANAGERS
• Han – 12pm

BAR
• Luke – 10am (KEY)
• Riley – 11am
• Parker – 3pm
• ME – 4pm (closer, KEY)
• Ben – 4pm

SERVERS
• Yoda – 12pm–8pm
• Rey – 3pm (closer)

HOSTS
• R2D2 – 11am
• Leia – 4pm

BARBACKS
• Anakin – 3pm

EVENTS / PARTIES
• 40@12 Football Fans
• 75@330 Company Happy Hour
• 100@615 Soccer Fans
```


## How it works

```
Gmail (read-only search)
   → .xlsx attachment
   → converted to a temporary Google Sheet (Drive API), read, then trashed
   → parsed into people × days × AM/PM cells
   → rules engine: end times + who's in charge
   → Google Calendar events (old events for that week replaced)
```

- **Parsing**: regular expressions turn cells like `4pmclky` into a start time plus flags (closer, key, server, bartender, training, volume, off). Sections are detected from the repeated day-header rows; barbacks are told apart from hosts by the blank row and gray fill between them.
- **Who's in charge**: the day is split at every manager/key arrival and departure, and each slice is assigned to the manager on duty or, if none, to the key hourlies present.
- **No duplicates**: events are tagged when created. When a week is processed again, tagged events for that week are deleted and recreated. Processed email IDs are stored so the timer doesn't redo them.

## Permissions

The script asks only for what it needs:

| Scope | Why |
|---|---|
| `gmail.readonly` | Find the schedule email and download the attachment. It never sends, labels, or deletes email. |
| `calendar` | Create and replace the shift events |
| `drive.file` | Only the temporary file it creates to read the Excel sheet |
| `spreadsheets` | Read that temporary sheet |
| `script.scriptapp` | The every-3-hours timer |


## Built with

JavaScript (Google Apps Script), Gmail API, Google Calendar, Google Drive API, regular expressions.

I designed the requirements and scheduling rules and tested the output against real weekly schedules; I used Claude for AI-assisted coding.

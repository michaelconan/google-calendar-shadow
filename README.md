# google-calendar-shadow
These Google Scripts are designed to create and update a "shadow" calendar that will copy events on your main calendar and send invitations to other accounts.

## Instructions:
- Make a copy of the Google Sheet in your own Drive.
    - Alternatively, import/convert Excel template and add code to script editor.
- Add Calendar ID if events are not to be shadowed from your personal calendar. To find ID, under 'My Calendars' on the Google Calendar page, select the 'Settings and Sharing' from the 'Options' menu, then select 'Integrate Calendar' and copy the Calendar ID listed.
- Add Emails to invite to shadow events in the "Calendars" tab
- Specify whether to copy full details (name / description, otherwise will be generic 'busy' label), and whether to shadow all events or only those accepted / owned
- Select 'Sync Calendar' from the 'Shadow' menu in the toolbar above and authorize, then select 'Sync Calendar' again
- [Optional] The script will result in a large number of email notifications being sent to the 'shadowing' email accounts. It is recommended that you create filters as indicated in the 'Filters' tab to keep your inbox clean!

## Steps Script will Take:
- Create shadow calendar if one has not been created by the script already
- Retrieve list of events for main and shadow calendars, apply filters based on specifications (remove 'free' events, filter on responses)
- If shadow event does not exist, create event, otherwise update event details and delete if cancelled
- Sets trigger to automatically perform incremental update to calendar events within the sync period (1 year)
- Sets trigger to automatically perform full sync every 30 days
- Displays summary message outlining results and process

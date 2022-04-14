/**
 * @module Calendar Shadow
 * @author Michael Conan
 * @description Script to create shadow events on alternate calendars (e.g. client accounts) that mirror main calendar,
 *  including all details or anonymized to 'busy' block
 * 
 */

// ~~~ GLOBALS ~~~
let SS = SpreadsheetApp.getActive();
let SHEET = SS.getSheetByName('Calendars');
const EMAILS = SHEET.getDataRange().getValues().map(r => r[1]).filter(v => v.includes('@'));
let MAINCAL = SHEET.getRange(4,4).getValue();
const DETAILS = SHEET.getRange(4,5).getValue();
const ACCEPT = SHEET.getRange(4,6).getValue();
let PROPS = PropertiesService.getUserProperties();
const MAXBACKOFF = 130;

/**
 * Function to create spreadsheet menu to start sync
 */
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Shadow')
  .addItem('Sync Calendar', 'fullSync')
  .addToUi();
}

/**
 * Function to clear out prior sync data and run full sync of non-updated events
 */
function fullSync() {
  
  // Delete sync token and run update
  PROPS.deleteProperty('syncToken');
  let stats = shadowCalendar();
  if (Object.keys(stats).includes('error')) {
    msg_('An error was encountered: ' + stats.error);
    return;
  }

  // Create event-based trigger if doesn't exist
  let func = arguments.callee.name;
  let triggers = ScriptApp.getProjectTriggers();
  if (!triggers.filter(t => t.getTriggerSource() == ScriptApp.TriggerSource.CLOCK && t.getHandlerFunction() == func).length) {
    ScriptApp.newTrigger(func)
    .timeBased()
    .everyDays(30)
    .create();
  }

  // Text table for logging
  let lengths = Object.keys(stats).map(s => String(s).length + 2);
  let borders = ['', ...lengths.map(l => '-'.repeat(l)), '\n'].join('+');
  let headers = ['', ...Object.keys(stats).map(s => ' ' + s + ' '), '\n'].join('|');
  let values = ['', ...Object.values(stats).map((s, i) => ' '.repeat(lengths[i] - String(s).length - 1) + s + ' '), '\n'].join('|');
  let result = ['', headers, values, ''].join(borders);
  Logger.log(result);

  // HTML table for message
  let html = '<table>' + [Object.keys(stats).map(s => '<th>'+s+'</th>').join(''), Object.values(stats).map(s => '<td>'+s+'</td>').join('')].map(r => '<tr>'+r+'</tr>').join('') + '</table>';
  let style = '<style>table {border-collapse: collapse;width:100%;}td,th {padding: 10px;border-bottom: 2px solid #8ebf42;text-align: center;}</style>'
  let message = '<p>Your shadow calendar has been created and invites have been sent to emails listed for events in the next year. A trigger has been set to update the shadow calendar as you update your primary calendar, and re-run a full sync monthly. To update prior events with new email addresses, run this sync function again.</p>';
  msg_(message + html + style, 'Sync Completed');
}

/**
 * Main function to run calendar shadow based on latest sync
 */
function shadowCalendar() {
  
  // Assign calendar to shadow
  if (MAINCAL == '') {
    MAINCAL = Calendar.Calendars.get('primary').id;
  } else {
    try {
      var test = Calendar.Calendars.get(MAINCAL);
    } catch (e) {
      let result = {
        error: 'Main calendar not found'
      }
      return result;
    }
  }
  Logger.log('Main calendar using: ' + MAINCAL);

  // Get shadow calendar from user properties
  let calendarId = PROPS.getProperty('shadow_calendar');
  
  // Create shadow calendar if none exists
  if (!calendarId) {
    let newCal = Calendar.Calendars.insert({summary: 'Shadow'});
    PROPS.setProperty('shadow_calendar', newCal.id);
    calendarId = newCal.id;
    Logger.log('Created new calendar with id: ' + calendarId);
  } else {
    Logger.log('Shadow calendar found with id: ' + calendarId);
  }
  let shadowCal = calendarId;

  // Get all main events
  var mainEvents = getCalendarEvents_(MAINCAL, true);
  Logger.log(mainEvents.length + ' total events...');

  // Remove events marked 'free'
  mainEvents = mainEvents.filter(e => e.transparency != 'transparent');
  Logger.log(mainEvents.length + ' events after transparent (free) removed...');

  // Remove events based on response and configuration
  mainEvents = mainEvents.filter(filterEventResponse_);
  Logger.log(mainEvents.length + ' events after response filter applied...');

  // Get list of all shadow calendar events
  let shadowEvents = getCalendarEvents_(shadowCal, false);
  Logger.log(shadowEvents.length + ' shadow events...');

  // Object for update stats
  let stats = {
    created: 0,
    updated: 0,
    deleted: 0
  };

  // Remove all-day events
  mainEvents = mainEvents.filter(e => e.start);
  mainEvents = mainEvents.sort((a, b) => a.start.dateTime - b.start.dateTime);

  // Create, update, delete shadow events
  stats = updateShadowEvents_(shadowCal, mainEvents, shadowEvents, stats);

  // Review / cleanup existing shadow events (dupes only unless full sync)
  stats = cleanupShadowEvents_(shadowCal, mainEvents, shadowEvents, stats, Boolean(PROPS.getProperty('syncToken')));
  
  // Create event-based trigger if doesn't exist
  let func = arguments.callee.name;
  let triggers = ScriptApp.getProjectTriggers();
  if (!triggers.filter(t => t.getTriggerSource() == ScriptApp.TriggerSource.CALENDAR && t.getHandlerFunction() == func).length) {
    ScriptApp.newTrigger(func)
    .forUserCalendar(MAINCAL)
    .onEventUpdated()
    .create();
  }
  
  return stats;
}

/**
 * Helper function to list all calendar events and store sync token if indicated
 */
function getCalendarEvents_(calId, sync) {
  let syncTok = PROPS.getProperty('syncToken');

  // Get 1 year of future events for main and shadow calendars
  let first = new Date();
  let last = new Date(new Date().setFullYear(first.getFullYear() + 1));
  let resource = {
    singleEvents: true,
  }
  if (sync && syncTok) {
    Logger.log('Initiating incremental sync...');
    resource.syncToken = syncTok;
  } else {
    if (sync)
      Logger.log('Initiating full sync...');
      
    resource.timeMin = first.toISOString();
    resource.timeMax = last.toISOString();
  }
  
  // Get list of all calendar events
  let allEvents = [];
  let pageToken = null;
  do {
    resource.pageToken = pageToken;
    var events = calendarEventCall_('list',[calId, resource]);
    allEvents = allEvents.concat(events.items);
    pageToken = events.nextPageToken;
  } while (pageToken);

  // Assign next sync token if incremental query
  if (sync && events.nextSyncToken) {
    Logger.log('Storing sync token for incremental updates...')
    PROPS.setProperty('syncToken', events.nextSyncToken);
  }

  return allEvents;
}

/**
 * Update shadow calendar to mirror main calendar events
 */
function updateShadowEvents_(shadowId, mainEvents, shadowEvents, result) {
  
  // Format emails as object for event
  let attds = EMAILS.map(e => Object({email: e}));

  // Add each event to shadow calendar based on configuration and add original ID as tag
  for (let evt of mainEvents) {
    
    // Event details
    let event = {
      attendees: attds,
      extendedProperties: {
        shared: {
          og: evt.id
        }
      }
    };
    if (evt.start) {
      event.start = evt.start;
      event.end = evt.end;
    } else {
      Logger.log('No start / end time specified...');
      continue;
    }
    if (DETAILS) {
        event.summary = evt.summary + '[shadow]';
        event.description = evt.description;
      } else {
        event.summary = 'busy [shadow]';
        event.description = 'shadow event';        
      }
    // Check for existing shadow events based on tag
    let existing = shadowEvents.filter(e => e.extendedProperties.shared.og == evt.id);
    if (!existing.length) {
      // No existing event, create new
      event = calendarEventCall_('insert', [event, shadowId, {sendUpdates: 'all'}]);
      Logger.log('Created shadow event for: ' + evt.summary);
      result.created += 1;
    } else {
      // Existing event, update or delete current
      Logger.log('Event already exists on shadow calendar: ' + evt.summary);
      let ogEvent = existing[0];

      // Delete if cancelled or newly declined based on user inputs
      if (evt.status == 'cancelled' || !filterEventResponse_(evt)) {
        // Delete (remove) cancelled/declined event
        event = calendarEventCall_('remove', [shadowId, ogEvent.id, {sendUpdates: 'all'}]);
        Logger.log('Deleted cancelled shadow event for: ' + evt.summary)
        result.deleted += 1;
      } else {
        // Compare details and update if needed

        // Check for differences in event metadata
        let diff = Object.keys(event).filter(k => {
          if (k == 'attendees') {
            // Compare email only
            return JSON.stringify(event[k].map(a => a.email)) != JSON.stringify(ogEvent[k].map(a => a.email));
          } else if (['start','end'].includes(k)) {
            // Standardize timezone / format to compare
            return new Date(Date.parse(event[k].dateTime)).toISOString() != new Date(Date.parse(ogEvent[k].dateTime)).toISOString();
          } else {
            // Compare all values
            return JSON.stringify(event[k]) != JSON.stringify(ogEvent[k]);
          }
        });
        
        // Update event if differences in metadata
        if (diff.length) {
          event = calendarEventCall_('patch', [event, shadowId, ogEvent.id, {sendUpdates: 'all'}]);
          Logger.log('Updated shadow event for: ' + evt.summary + '\nUpdated fields: ' + diff.join(', '));
          result.updated += 1;
        } else {
          Logger.log('No updates to: ' + evt.summary);
        }
      }
    }
  }

  return result;
}

/**
 * Helper to remove old or duplicated shadow events
 */
function cleanupShadowEvents_(shadowId, mainEvents, shadowEvents, result, dupeOnly) {
  
  // Cleanup shadow calendar from old / duplicate events
  let mainIds = mainEvents.map(e => e.id);
  let shadowIds = shadowEvents.map(e => e.extendedProperties.shared.og);

  // Events where OG ID not in mainevents or not first instance of OG ID
  // duplicates?  || shadowIds.indexOf(e.extendedProperties.shared.og) != i
  let oldShadowEvents = shadowEvents.map((e, i, arr) => {
    if (shadowIds.indexOf(e.extendedProperties.shared.og) != i) {
      e.shadowCategory = 'duplicate';
    } else if (!mainIds.includes(e.extendedProperties.shared.og)) {
      e.shadowCategory = 'missing';
    } else {
      e.shadowCategory = 'found';
    }
    return e;
  });
  oldShadowEvents = oldShadowEvents.filter(e => e.shadowCategory != 'found');
  if (dupeOnly) {
    oldShadowEvents = oldShadowEvents.filter(e => e.shadowCategory != 'missing');
  }
  Logger.log(oldShadowEvents.length + ' old events on shadow calendar no longer on main...');
  for (let ogEvent of oldShadowEvents) {
    //event = calendarEventCall_('remove', [shadowId, ogEvent.id, {sendUpdates: 'all'}]);
    Logger.log('Deleted old shadow event for: ' + ogEvent.summary + ' identified as: ' + ogEvent.shadowCategory + ' at: ' + ogEvent.start.dateTime + ' with original id: ' + ogEvent.extendedProperties.shared.og);
    result.deleted += 1;
  }
}

/**
 * Function to implement exponential backoff of API calls
 */
function calendarEventCall_(method, args) {
  let backoff = 1;
  let result;
  let err;
  // Attempt specified call with incremental wait time until hit max, then error
  do {   
    try {
      result = Calendar.Events[method](...args);   
      return result;
    } catch (e) {
      err = e;
      Logger.log('backoff: ' + backoff + ' - ' + method + ' - ' + JSON.stringify(args));
      Logger.log(e);
      Utilities.sleep(backoff * 1000);
      backoff *= 2;
    }
  } while (backoff <= MAXBACKOFF);
  if (err.message.includes('Calendar usage limits exceeded.')) {
    err.message = 'Error may relate to too many notifications to external domains, or too many requests in short period of time -- ' + err.message;
  }
  throw err;
}

/**
 * Helper function to apply user-defined logic for response filtering
 */
function filterEventResponse_(event) {
  // Check event has attendees
  if (event.attendees) {
    // Check self as attendee
    let self = event.attendees.filter(a => a.self);
    if (self.length) {
      if (self[0].responseStatus == 'declined') {
        return false;
      } else if (self[0].responseStatus == 'accepted') {
        return true;
      }
      return !ACCEPT; // False if only accepted events, otherwise true (tentative / need action)
    }
  }
  return true;
}

// Modeless message box
function msg_(text, title) {
  var msgBox = HtmlService
    .createHtmlOutput(text)
    .setWidth(450)
    .setHeight(250);
  try {
    SpreadsheetApp.getUi()
    .showModelessDialog(msgBox, title);
  } catch (e) {
  }
}

function test() {
  let e = Calendar.Events.get('michael.conan@pwc.com','skb3rjv21kqbap108se6dltv4t_20220328T153000Z');
  Logger.log(Boolean(null));
  Logger.log(Boolean('abcd'));

}
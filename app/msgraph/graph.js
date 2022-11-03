const graph = require('@microsoft/microsoft-graph-client');
const config = require('../../config/config');
const moment = require('moment');
require('isomorphic-fetch');

// TODO: Move this somewhere else?
Date.prototype.addDays = function (days) {
  const date = new Date(this.valueOf());
  date.setDate(date.getDate() + days);
  return date;
};

module.exports = {
  getRoomList: async (msalClient) => {
    const client = getAuthenticatedClient(msalClient);

    const roomlist = await client
      .api('/places/microsoft.graph.roomlist')
      .select('displayName, emailAddress')
      .orderby('displayName')
      .top(parseInt(config.calendarSearch.maxRoomLists))
      .get();
    return roomlist;
  },

  getRooms: async (msalClient, email) => {
    const client = getAuthenticatedClient(msalClient);

    const rooms = await client
      .api(`/places/${email}/microsoft.graph.roomlist/rooms`)
      .select('displayName,emailAddress')
      .orderby('displayName')
      .top(parseInt(config.calendarSearch.maxRooms))
      .get();
    return rooms;
  },

  bookRoom: async (msalClient, roomEmail, roomName, start, end, bookingType, subject, body) => {
    const client = getAuthenticatedClient(msalClient);
    if ((bookingType === 'BookNow') || (bookingType === 'BookAfter')) {
      const event = {
        subject: subject,
        body: {
          contentType: 'HTML',
          content: body
        },
        start: {
          dateTime: start,
          timeZone: 'UTC'
        },
        end: {
          dateTime: end,
          timeZone: 'UTC'
        },
        location: {
          displayName: roomEmail
        },
        attendees: [
          {
            type: 'required',
            emailAddress: {
              address: roomEmail
            }
          }
        ]
      };

      const response = await client
        .api(`/users/${roomEmail}/calendar/events`)
        .post(event);
      return response;

      // TODO - Review this code that is mostly added by copilot and not tested 🙈
    } else if ((bookingType === 'Extend') || (bookingType === 'EndNow')) {
      // Get the events for the room from now and to x days/years in the future
      const queryNowDateTime = new Date();
      const queryEndDateTime = new Date(now.getTime() + + 576000000);

      const response = await client
        .api(`/users/${roomEmail}/calendar/calendarView?startDateTime=${queryNowDateTime.toISOString()}&endDateTime=${queryEndDateTime.toISOString()}`)
        .select('subject,start,end')
        .orderby('Start/DateTime')
        .top(parseInt(config.calendarSearch.maxItems))
        .get();

      // Find the meeting that is currently running
      const currentMeeting = response.value.find((meeting) => {
        const start = new Date(meeting.start.dateTime);
        const end = new Date(meeting.end.dateTime);
        return (start < now) && (now < end);
      }
      );

      if (currentMeeting) {
        // Extend the meeting
        const newEnd = new Date(end);
        currentMeeting.end.dateTime = newEnd.toISOString();
        currentMeeting.end.timeZone = 'UTC';

        const response = await client
          .api(`/users/${roomEmail}/calendar/events/${currentMeeting.id}`)
          .patch(currentMeeting);
        return response;
      }

      // No meeting found
      return null;

    } else {
      throw new Error(`Invalid booking type: ${bookingType}`);
    }
  },

  getCalendarView: async (msalClient, email) => {
    const client = getAuthenticatedClient(msalClient);

    const start_datetime = new Date();
    const end_datetime = start_datetime.addDays(parseInt(config.calendarSearch.maxDays));

    const events = await client
      .api(`/users/${email}/calendar/calendarView`)
      .query(`startDateTime=${start_datetime.toISOString()}&endDateTime=${end_datetime.toISOString()}`)
      .select('organizer,subject,start,end,sensitivity')
      .orderby('Start/DateTime')
      .top(parseInt(config.calendarSearch.maxItems))
      .get();
    return events;
  }
};

function getAuthenticatedClient(msalClient) {
  if (!msalClient) {
    throw new Error(`Invalid MSAL state. Client: ${msalClient ? 'present' : 'missing'}`);
  }

  const clientCredentialRequest = {
    scopes: ['.default'],
    skipCache: true // (optional) this skips the cache and forces MSAL to get a new token from Azure AD
  };

  // Initialize Graph client
  const client = graph.Client.init({
    // Implement an auth provider that gets a token
    // from the app's MSAL instance
    authProvider: async (done) => {
      try {
        const response = await msalClient.acquireTokenByClientCredential(clientCredentialRequest);

        // First param to callback is the error,
        // Set to null in success case
        done(null, response.accessToken);
      } catch (err) {
        console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
        done(err, null);
      }
    }
  });

  return client;
}

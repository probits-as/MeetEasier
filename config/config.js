const msal = require('@azure/msal-node');
const cacheLocation = './data/cache.json';
const cachePlugin = require('../app/msgraph/cachePlugin')(cacheLocation);

// Read the .env-file
require('dotenv').config();

// expose our config directly to our application using module.exports
module.exports = {
  // Configuration for the msgraph-sdk
  msalConfig: {
    auth: {
      clientId: process.env.OAUTH_CLIENT_ID ? process.env.OAUTH_CLIENT_ID : 'OAUTH_CLIENT_ID_NOT_SET',
      authority: process.env.OAUTH_AUTHORITY ? process.env.OAUTH_AUTHORITY : 'OAUTH_AUTHORITY_NOT_SET',
      clientSecret: process.env.OAUTH_CLIENT_SECRET
        ? process.env.OAUTH_CLIENT_SECRET
        : 'OAUTH_CLIENT_SECRET_NOT_SET'
    },
    /* 		cache: {
          cachePlugin
        }, */
    system: {
      loggerOptions: {
        // TODO: Better logging / other options?
        loggerCallback(loglevel, message, containsPii) {
          console.log(message);
        },
        piiLoggingEnabled: false,
        logLevel: msal.LogLevel.Error
      }
    }
  },

  // Search-settings to use when retrieving data from the calendars
  calendarSearch: {
    maxDays: process.env.SEARCH_MAXDAYS ? process.env.SEARCH_MAXDAYS : 10,
    maxRoomLists: process.env.SEARCH_MAXROOMLISTS ? process.env.SEARCH_MAXROOMLISTS : 10,
    maxRooms: process.env.SEARCH_MAXROOMS ? process.env.SEARCH_MAXROOMS : 10,
    maxItems: process.env.SEARCH_MAXITEMS ? process.env.SEARCH_MAXITEMS : 6
  }
};

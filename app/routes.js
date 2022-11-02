const msal = require('@azure/msal-node');
const config = require('../config/config');

const msalClient = new msal.ConfidentialClientApplication(config.msalConfig);

module.exports = function (app) {
  var path = require('path');

  // api routes ================================================================
  // returns an array of room objects
  app.get('/api/rooms', function (req, res) {
    let api;
    if (config.calendarSearch.useGraphAPI === 'true') {
      api = require('./msgraph/rooms.js');
    } else {
      api = require('./ews/rooms.js');
    }

    api(function (err, rooms) {
      if (err) {
        if (err.responseCode === 127) {
          res.json({
            error:
              'Oops, there seems to be an issue with the credentials you have supplied.  Make sure you typed them correctly and that you have access to Exchange Roomlists.'
          });
        } else {
          res.json({
            error: 'Hmm, there seems to be a weird issue occuring.'
          });
        }
      } else {
        res.json(rooms);
      }
    }, msalClient);
  });

  // returns an array of roomlist objects
  app.get('/api/roomlists', function (req, res) {
    let api;
    if (config.calendarSearch.useGraphAPI === 'true') {
      api = require('./msgraph/roomlists.js');
    } else {
      api = require('./ews/roomlists.js');
    }

    api(function (err, roomlists) {
      if (err) {
        if (err.responseCode === 127) {
          res.json({
            error:
              'Oops, there seems to be an issue with the credentials you have supplied.  Make sure you typed them correctly and that you have access to Exchange Roomlists.'
          });
        } else {
          res.json({
            error: 'Hmm, there seems to be a weird issue occuring.'
          });
        }
      } else {
        res.json(roomlists);
      }
    }, msalClient);
  });

  // books a room
  app.get('/api/roombooking', function (req, res) {
    let api;
    if (config.calendarSearch.useGraphAPI === 'true') {
      api = require('./msgraph/roombooking.js');
    } else {
      api = require('./ews/roombooking.js');
    }

    console.log("Route Room Booking");

    var roomEmail = req.query.roomEmail;
    var roomName = req.query.roomName;
    var startTime = req.query.startTime;
    var endTime = req.query.endTime;

    api(function (err) {
      if (err) {
        if (err.responseCode === 127) {
          res.json({
            error:
              'Oops, there seems to be an issue with the credentials you have supplied.  Make sure you typed them correctly and that you have access to Exchange Roomlists.'
          });
        } else {
          res.json({
            error: 'Hmm, there seems to be a weird issue occuring.'
          });
        }
      } else {
        res.json({
          success: 'Room booked successfully!'
        });
      }
    }, roomEmail, roomName, startTime, endTime, null, msalClient);

    console.log(roomEmail + " | " + roomName + " | " + startTime + " | " + endTime);
  });

  // heartbeat-service to check if server is alive
  app.get('/api/heartbeat', function (req, res) {
    res.json({ status: 'OK' });
  });

  // redirects everything else to our react app
  app.get('*', function (req, res) {
    res.sendFile(path.join(__dirname, '../ui-react/build/', 'index.html'));
  });
};

const graph = require('./graph');
const blacklist = require('../../config/room-blacklist.js');

module.exports = function (callback, msalClient) {
  // promise: get all room lists
  const getListOfRooms = () => {
    var promise = new Promise(function (resolve, reject) {
      graph
        .getRoomList(msalClient)
        .then(
          (lists) => {
            var roomLists = lists.value;
            resolve(roomLists);
          },
          (err) => {
            callback(err, null);
          }
        )
        .catch((err) => callback(err, null));
    });

    return promise;
  };

  // promise: get all rooms in room lists
  const getRoomsInLists = (roomLists) => {
    var promise = new Promise(function (resolve, reject) {
      let roomAddresses = [];
      let counter = 0;

      roomLists.forEach(function (item, i, array) {
        graph.getRooms(msalClient, item.emailAddress).then((rooms) => {
          rooms.value.forEach(function (roomItem, roomIndex, roomsArray) {
            // use either email var or roomItem.Address - depending on your use case
            let inBlacklist = isRoomInBlacklist(roomItem.emailAddress);

            // if not in blacklist, proceed as normal; otherwise, skip
            if (!inBlacklist) {
              let room = {};
              let email = roomItem.emailAddress;
              let roomAlias = roomItem.displayName.toLowerCase().replace(/\s+/g, '-');

              room.Roomlist = item.displayName;
              room.Name = roomItem.displayName;
              room.RoomAlias = roomAlias;
              room.Email = email;
              roomAddresses.push(room);
            }
          });
          counter++;

          if (counter === array.length) {
            resolve(roomAddresses);
          }
        });
      });
    });
    return promise;
  };

  const fillRoomData = (context, room, appointments = [], option = {}) => {
    room.Appointments = [];
    appointments.forEach(function (appt, index) {
      // get start time from appointment
      const start = processTime(appt.start.dateTime),
        end = processTime(appt.end.dateTime),
        now = Date.now();

      room.Busy = index === 0 ? start < now && now < end : room.Busy;

      let isAppointmentPrivate = appt.sensitivity.toLowerCase() === 'normal' ? false : true;
      let subject = isAppointmentPrivate ? 'Private' : appt.subject;

      room.Appointments.push({
        Subject: subject,
        Organizer: appt.organizer.emailAddress.name,
        Start: start,
        End: end,
        Private: isAppointmentPrivate
      });
    });

    if (option.errorMessage) room.ErrorMessage = option.errorMessage;
    context.itemsProcessed++;

    if (context.itemsProcessed === context.roomAddresses.length) {
      context.roomAddresses.sort((a, b) => a.Name.toLowerCase().localeCompare(b.Name.toLowerCase()));
      context.callback(context.roomAddresses);
    }
  };

  // promise: get current or upcoming appointments for each room
  const getAppointmentsForRooms = (roomAddresses) => {
    var promise = new Promise(function (resolve, reject) {
      var context = {
        callback: resolve,
        itemsProcessed: 0,
        roomAddresses
      };

      roomAddresses.forEach(function (room, index, array) {
        graph.getCalendarView(msalClient, room.Email).then(
          (response) => {
            fillRoomData(context, room, response.value);
          },
          (error) => {
            // handle the error here
            // callback(error, null);
            fillRoomData(context, room, undefined, { errorMessage: error });
          }
        );
      });
    });
    return promise;
  };

  // check if room is in blacklist
  const isRoomInBlacklist = (email) => blacklist.roomEmails.includes(email);

  // do all of the process for the appointment times
  // process the time from UTC to localtime
  const processTime = (appointmentTime) => {
    const date = new Date(appointmentTime);
    const localOffset = -1 * date.getTimezoneOffset() * 60000;
    const timestamp = date.getTime() + localOffset;

    return timestamp;
  };

  // perform promise chain to get rooms
  getListOfRooms().then(getRoomsInLists).then(getAppointmentsForRooms).then((rooms) => callback(null, rooms));
};

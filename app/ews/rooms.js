module.exports = function(callback) {
	// modules -------------------------------------------------------------------
	var ews = require('ews-javascript-api');
	var config = require('../../config/config');
	var blacklist = require('../../config/room-blacklist.js');

	// ews -----------------------------------------------------------------------
	// - TODO: Make the exchangeserver-version configurable
	var exch = new ews.ExchangeService(ews.ExchangeVersion.Exchange2016);
	exch.Credentials = new ews.ExchangeCredentials(config.exchange.username, config.exchange.password);
	exch.Url = new ews.Uri(config.exchange.uri);

	// promise: get all room lists
	var getListOfRooms = function() {
		var promise = new Promise(function(resolve, reject) {
			exch.GetRoomLists().then(
				(lists) => {
					var roomLists = lists.items;
					resolve(roomLists);
				},
				(err) => {
					callback(err, null);
				}
			);
		});
		return promise;
	};

	// promise: get all rooms in room lists
	var getRoomsInLists = function(roomLists) {
		var promise = new Promise(function(resolve, reject) {
			var roomAddresses = [];
			var counter = 0;

			roomLists.forEach(function(item, i, array) {
				exch.GetRooms(new ews.Mailbox(item.Address)).then((rooms) => {
					rooms.forEach(function(roomItem, roomIndex, roomsArray) {
						// use either email var or roomItem.Address - depending on your use case
						let inBlacklist = isRoomInBlacklist(roomItem.Address);

						// if not in blacklist, proceed as normal; otherwise, skip
						if (!inBlacklist) {
							let room = {};

							// if the email addresses != your corporate domain,
							// replace email domain with domain
							let email = roomItem.Address;
							email = email.substring(0, email.indexOf('@'));
							email = email + '@' + config.domain;

							let roomAlias = roomItem.Name.toLowerCase().replace(/\s+/g, '-');

							room.Roomlist = item.Name;
							room.Name = roomItem.Name;
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

	var fillRoomData = function(context, room, appointments = [], option = {}) {
		room.Appointments = [];
		appointments.forEach(function(appt, index) {
			// get start time from appointment
			var start = processTime(appt.Start.momentDate),
				end = processTime(appt.End.momentDate),
				now = Date.now();

			room.Busy = index === 0 ? start < now && now < end : room.Busy;

			let isAppointmentPrivate = appt.Sensitivity === 'Normal' ? false : true;

			let subject = isAppointmentPrivate ? 'Private' : appt.Subject;

			room.Appointments.push({
				Subject: subject,
				Organizer: appt.Organizer.Name,
				Start: start,
				End: end,
				Private: isAppointmentPrivate
			});
		});

		if (option.errorMessage) {
			room.ErrorMessage = option.errorMessage;
		}

		context.itemsProcessed++;

		if (context.itemsProcessed === context.roomAddresses.length) {
			context.roomAddresses.sort((a, b) => a.Name.toLowerCase().localeCompare(b.Name.toLowerCase()));
			context.callback(context.roomAddresses);
		}
	};

	// promise: get current or upcoming appointments for each room
	var getAppointmentsForRooms = function(roomAddresses) {
		var promise = new Promise(function(resolve, reject) {
			var context = {
				callback: resolve,
				itemsProcessed: 0,
				roomAddresses
			};

			roomAddresses.forEach(function(room, index, array) {
				var calendarFolderId = new ews.FolderId(ews.WellKnownFolderName.Calendar, new ews.Mailbox(room.Email));
				var view = new ews.CalendarView(ews.DateTime.Now, ews.DateTime.Now.AddDays(10), 6);
				exch.FindAppointments(calendarFolderId, view).then(
					(response) => {
						fillRoomData(context, room, response.Items);
					},
					(error) => {
						// handle the error here
						// callback(error, null);
						fillRoomData(context, room, undefined, { errorMessage: error.response.errorMessage });
					}
				);
			});
		});
		return promise;
	};

	// check if room is in blacklist
	function isRoomInBlacklist(email) {
		return blacklist.roomEmails.includes(email);
	}

	// do all of the process for the appointment times
	function processTime(appointmentTime) {
		var time = JSON.stringify(appointmentTime);
		time = time.replace(/"/g, '');
		var time = new Date(time);
		var time = time.getTime();

		return time;
	}

	// perform promise chain to get rooms
	getListOfRooms().then(getRoomsInLists).then(getAppointmentsForRooms).then(function(rooms) {
		callback(null, rooms);
	});
};

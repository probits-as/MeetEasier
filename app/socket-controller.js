const msal = require('@azure/msal-node');
const config = require('../config/config');

const msalClient = new msal.ConfidentialClientApplication(config.msalConfig);

module.exports = function(io) {
	var isRunning = false;

	// Check and update rooms every 60 seconds -----------------------------------
	io.of('/').on('connection', function(socket) {
		if (!isRunning) {
			(function callAPI() {
				let api;
				api = require('./msgraph/rooms.js');

				api(function(err, result) {
					if (result) {
						if (err) return console.log(err);
						// send data to page
						io.of('/').emit('updatedRooms', result);
					}

					io.of('/').emit('controllerDone', 'done');
				}, msalClient);

				setTimeout(callAPI, 60000);
			})();
		}

		isRunning = true;
	});
};

import React from 'react';
import PropTypes from 'prop-types';
import { Link } from 'react-router-dom';
import * as moment from 'moment';
import ReactDOM from 'react-dom';

class FFButton extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      isVisible: false,
      count: 0,
      previousElement: null
    };
  }
  componentWillMount() {
    document.addEventListener('mousedown', this.handleClickOutside, false);
  }

  // Unbind event on unmount to prevent leaks
  componentWillUnmount() {
    window.removeEventListener('mousedown', this.handleClickOutside, false);
  }

  handleClickOutside = (event) => {

    if (this.state.previousElement != null) {
      if (event.target.className.toString().indexOf(this.state.previousElement.toString())) {
        this.setState((state) => {
          return { isVisible: false }
        });
      }
      else {
      }
    }
    this.setState((state) => {
      return { previousElement: event.target.className }
    });
  }
  incrementCount() {
    this.setState((state) => {
      return { count: state.count + 1 }
    });
  }
  setVisible() {
    this.setState((state) => {
      return { isVisible: !state.isVisible }
    });
  }
  _handleClick = (e) => {
    e.preventDefault();
    this.incrementCount();
    this.setVisible();
  }
  renderDropdown() {
    const room = this.props.room;
    let currentAppointment;
    let currentAppointmentEnd;
    let currentAppointmentStart;
    let nextAppointment;
    let nextAppointmentEnd;
    let nextAppointmentStart;
    let showPopup = this.props.showPopup;
    if (room.Appointments[0]) {
      currentAppointment = room.Appointments[0];
      currentAppointmentEnd = new Date(parseInt(room.Appointments[0].End, 10));
      currentAppointmentStart = new Date(parseInt(room.Appointments[0].Start, 10));
    }
    let contentDropdown;
    if (this.props.BtnFunc == "BookAfter") {
      contentDropdown = this.props.DropdownContent.map((number) =>
        <button type="button" class="btn BookAfter btn-drop " disabled={showPopup} onClick={(e) => { BookAfter(number, room, currentAppointmentEnd, this.props.togglePopup); this.setVisible() }}>{number} Minutes</button>
      );
    }
    else if (this.props.BtnFunc == "ExtendBooking") {
      contentDropdown = this.props.DropdownContent.map((number) =>
        <button type="button" class="btn ExtendBooking btn-drop " disabled={showPopup} onClick={(e) => { ExtendBooking(number, room, currentAppointmentStart, currentAppointmentEnd, this.props.togglePopup); this.setVisible() }}>{number} Minutes</button>
      );
    }
    else if (this.props.BtnFunc == "BookNow") {
      contentDropdown = this.props.DropdownContent.map((number) =>
        <button type="button" class="btn BookNow btn-drop" disabled={showPopup} onClick={(e) => { BookNow(number, room, this.props.togglePopup); this.setVisible() }}>{number} Minutes</button>
      );
    }
    else if (this.props.BtnFunc == "BookAfterNext") {
      contentDropdown = this.props.DropdownContent.map((number) =>
        <button type="button" class="btn BookAfterNext btn-drop" disabled={showPopup} onClick={(e) => { BookAfter(number, room, currentAppointmentEnd, this.props.togglePopup); this.setVisible() }}>{number} Minutes</button>
      );
    }
    else if (this.props.BtnFunc == "EndNow") {
      contentDropdown = <div>
        <button type="button" class="btn EndNow btn-drop" disabled={showPopup}>Please Confirm:</button>
        <button type="button" class="btn EndNow btn-drop" disabled={showPopup} onClick={e => { EndNow(room, currentAppointmentStart, currentAppointmentEnd, this.props.togglePopup); this.setVisible() }}>YES</button>
        <button type="button" class="btn EndNow btn-drop" disabled={showPopup} onClick={e => { this.setVisible() }}>NO</button>
      </div>
    }
    else {
      contentDropdown = "";
    }
    return (
      <ul className="dropdown-content1">
        {contentDropdown}
      </ul>
    );
  }
  render() {
    let classString = "btn ";
    classString = classString + this.props.BtnFunc;
    return (
      <div>
        <div class="dropdown">
          <button class={classString} type="button" disabled={this.props.showPopup} onClick={(e) => this._handleClick(e)} tabindex="1" >{this.props.ButtonTitle}
          </button>
          {this.state.isVisible ? this.renderDropdown() : null}
        </div>
      </div>
    );
  }
}
const Details = ({ room, details }) => (
  <div id="single-room__details">
    {details.appointmentExists &&
      <div id="single-room__meeting-title">
        <span id="single-room__next-up">
          {details.nextUp}
        </span>
        <span id="single-room__meeting-subject">
          {room.Appointments[0].Subject}
        </span>
      </div>
    }
  </div>
);

const Time = ({ room, details }) => (
  <div id="single-room__meeting-time">
    {details.appointmentExists &&
      new Date(parseInt(room.Appointments[0].Start, 10)).toLocaleTimeString([], { weekday: 'short', hour: '2-digit', minute: '2-digit' })
      + ' - ' +
      new Date(parseInt(room.Appointments[0].End, 10)).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
    }
  </div>
);

const Organizer = ({ room, details }) => {
  return (
    <div id="single-room__meeting-organizer">
      {details.appointmentExists &&
        room.Appointments[0].Organizer
      }
    </div>
  );
};
function BookAfter(time, room, startTimeDT, togglePopup) {
  togglePopup("Booking Now... Please Wait!");
  var startTime = moment(startTimeDT).add(1, 'minutes').toISOString();
  var endTime = moment(startTime).add(time, 'minutes').toISOString();
  var roomEmail = room.Email;
  var roomName = room.Name;
  var bookingType = "BookAfter";
  var data = { roomEmail: roomEmail, roomName: roomName, startTime: startTime, endTime: endTime };
  fetch(`../api/roombooking?roomEmail=${encodeURIComponent(data.roomEmail)}&roomName=${encodeURIComponent(data.roomName)}&startTime=${encodeURIComponent(data.startTime)}&endTime=${encodeURIComponent(data.endTime)}&bookingType=${encodeURIComponent(bookingType)}`);
  setTimeout(reloadPage, 5000);
}
function BookNow(time, room, togglePopup) {
  togglePopup("Booking Now... Please Wait!");
  var startTimeDT = new Date();
  var endTimeDT = new Date(startTimeDT.getTime() + time * 60000);
  var startTime = moment(startTimeDT).toISOString();
  var endTime = moment(endTimeDT).toISOString();
  var roomEmail = room.Email;
  var roomName = room.Name;
  var bookingType = "BookNow";
  var data = { roomEmail: roomEmail, roomName: roomName, startTime: startTime, endTime: endTime };
  fetch(`../api/roombooking?roomEmail=${encodeURIComponent(data.roomEmail)}&roomName=${encodeURIComponent(data.roomName)}&startTime=${encodeURIComponent(data.startTime)}&endTime=${encodeURIComponent(data.endTime)}&bookingType=${encodeURIComponent(bookingType)}`);
  setTimeout(reloadPage, 5000);
}
function ExtendBooking(time, room, startTimeDT, endTimeDT, togglePopup) {
  togglePopup("Extending Now... Please Wait!");
  var bookingType = "Extend";
  endTimeDT = new Date(endTimeDT.getTime() + time * 60000);
  var startTime = moment(startTimeDT).toISOString();
  var endTime = moment(endTimeDT).toISOString();
  var roomEmail = room.Email;
  var roomName = room.Name;
  var data = { roomEmail: roomEmail, roomName: roomName, startTime: startTime, endTime: endTime };
  fetch(`../api/roombooking?roomEmail=${encodeURIComponent(data.roomEmail)}&roomName=${encodeURIComponent(data.roomName)}&startTime=${encodeURIComponent(data.startTime)}&endTime=${encodeURIComponent(data.endTime)}&bookingType=${encodeURIComponent(bookingType)}`);
  setTimeout(reloadPage, 5000);
}
function EndNow(room, startTimeDT, endTimeDT, togglePopup) {
  togglePopup("Terminating Booking... Please Wait");
  var bookingType = "EndNow";
  endTimeDT = new Date();
  var startTime = moment(startTimeDT).toISOString();
  var endTime = moment(endTimeDT).toISOString();
  var roomEmail = room.Email;
  var roomName = room.Name;
  var data = { roomEmail: roomEmail, roomName: roomName, startTime: startTime, endTime: endTime };
  fetch(`../api/roombooking?roomEmail=${encodeURIComponent(data.roomEmail)}&roomName=${encodeURIComponent(data.roomName)}&startTime=${encodeURIComponent(data.startTime)}&endTime=${encodeURIComponent(data.endTime)}&bookingType=${encodeURIComponent(bookingType)}`);
  setTimeout(reloadPage, 5000);
}
function reloadPage() {
  window.location.reload();
}
function ButtonControl(props) {
  let room = props.room;
  let showPopup = props.showPopup;
  let details = props.details;
  var moment = require('moment');
  if (room.Busy) {
    let currentAppointment = room.Appointments[0];
    let currentAppointmentEnd = new Date(parseInt(currentAppointment.End, 10));
    let currentAppointmentStart = new Date(parseInt(currentAppointment.Start, 10));
    if (room.Appointments.length > 1) {
      let nextAppointmentStart = new Date(parseInt(room.Appointments[1].Start, 10));
      currentAppointmentEnd = moment(currentAppointmentEnd);
      nextAppointmentStart = moment(nextAppointmentStart);
      let timeDifference = nextAppointmentStart.diff(currentAppointmentEnd, 'minutes')
      if (timeDifference > 120) {
        let DropdownContent = [15, 30, 60, 120]
        return (
          <div> <h4> Extend/Book after this meeting Now: </h4>
            <table class="buttonContainer"><tr>
              <td class="td-btn"><FFButton ButtonTitle="Book After" DropdownContent={DropdownContent} BtnFunc="BookAfter" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="Extend" DropdownContent={DropdownContent} BtnFunc="ExtendBooking" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="End Meeting" DropdownContent={DropdownContent} BtnFunc="EndNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
            </tr></table></div>
        );
      }
      else if (timeDifference > 90) {
        let DropdownContent = [15, 30, 60, 90]
        return (
          <div> <h4> Extend/Book after this meeting Now: </h4>
            <table class="buttonContainer"><tr>
              <td class="td-btn"><FFButton ButtonTitle="Book After" DropdownContent={DropdownContent} BtnFunc="BookAfter" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="Extend" DropdownContent={DropdownContent} BtnFunc="ExtendBooking" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="End Meeting" DropdownContent={DropdownContent} BtnFunc="EndNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
            </tr></table></div>
        );
      }
      else if (timeDifference > 60) {
        let DropdownContent = [15, 30, 45, 60]
        return (
          <div> <h4> Extend/Book after this meeting Now: </h4>
            <table class="buttonContainer"><tr>
              <td class="td-btn"><FFButton ButtonTitle="Book After" DropdownContent={DropdownContent} BtnFunc="BookAfter" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="Extend" DropdownContent={DropdownContent} BtnFunc="ExtendBooking" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="End Meeting" DropdownContent={DropdownContent} BtnFunc="EndNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
            </tr></table></div>
        );
      }
      else if (timeDifference > 45) {
        let DropdownContent = [15, 30, 45]
        return (
          <div> <h4> Extend/Book after this meeting Now: </h4>
            <table class="buttonContainer"><tr>
              <td class="td-btn"><FFButton ButtonTitle="Book After" DropdownContent={DropdownContent} BtnFunc="BookAfter" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="Extend" DropdownContent={DropdownContent} BtnFunc="ExtendBooking" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="End Meeting" DropdownContent={DropdownContent} BtnFunc="EndNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
            </tr></table></div>
        );
      }
      else if (timeDifference > 30) {
        let DropdownContent = [15, 30]
        return (
          <div> <h4> Extend/Book after this meeting Now: </h4>
            <table class="buttonContainer"><tr>
              <td class="td-btn"><FFButton ButtonTitle="Book After" DropdownContent={DropdownContent} BtnFunc="BookAfter" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="Extend" DropdownContent={DropdownContent} BtnFunc="ExtendBooking" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="End Meeting" DropdownContent={DropdownContent} BtnFunc="EndNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
            </tr></table></div>
        );
      }
      else if (timeDifference > 15) {
        let DropdownContent = [15, timeDifference]
        return (
          <div> <h4> Extend/Book after this meeting Now: </h4>
            <table class="buttonContainer"><tr>
              <td class="td-btn"><FFButton ButtonTitle="Book After" DropdownContent={DropdownContent} BtnFunc="BookAfter" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="Extend" DropdownContent={DropdownContent} BtnFunc="ExtendBooking" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="End Meeting" DropdownContent={DropdownContent} BtnFunc="EndNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
            </tr></table></div>
        );
      }
      else if (timeDifference >= 5) {
        let DropdownContent = [timeDifference]
        return (
          <div> <h4> Extend/Book after this meeting Now: </h4>
            <table class="buttonContainer"><tr>
              <td class="td-btn"><FFButton ButtonTitle="Book After" DropdownContent={DropdownContent} BtnFunc="BookAfter" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="Extend" DropdownContent={DropdownContent} BtnFunc="ExtendBooking" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
              <td class="td-btn"><FFButton ButtonTitle="End Meeting" DropdownContent={DropdownContent} BtnFunc="EndNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
            </tr></table></div>
        );
      }
      else {
        let DropdownContent = [0]
        return (<div> <h4> Room is unavailable </h4>
          <div class="dropdown">
            <FFButton ButtonTitle="End Meeting" DropdownContent={DropdownContent} BtnFunc="EndNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
          </div>
        </div>);
      }
    }
    else {
      let DropdownContent = [15, 30, 60, 120]
      return (
        <div> <h4> Extend/Book after this meeting Now: </h4>
          <table class="buttonContainer"><tr>
            <td class="td-btn"><FFButton ButtonTitle="Book After" DropdownContent={DropdownContent} BtnFunc="BookAfter" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
            <td class="td-btn"><FFButton ButtonTitle="Extend" DropdownContent={DropdownContent} BtnFunc="ExtendBooking" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
            <td class="td-btn"><FFButton ButtonTitle="End Meeting" DropdownContent={DropdownContent} BtnFunc="EndNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} /></td>
          </tr></table></div>
      );
    }
  }
  else {
    if (room.Appointments.length > 0) {
      let now = new Date();
      let nextAppointmentStart = new Date(parseInt(room.Appointments[0].Start, 10));
      now = moment(now);
      nextAppointmentStart = moment(nextAppointmentStart);
      let timeDifference = nextAppointmentStart.diff(now, 'minutes')
      if (timeDifference > 120) {
        let DropdownContent = [15, 30, 60, 120];
        return (
          <div> <h4> Book This Room Now: </h4>
            <FFButton ButtonTitle="Book Now" DropdownContent={DropdownContent} BtnFunc="BookNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
          </div>
        );
      }
      else if (timeDifference > 90) {
        let DropdownContent = [15, 30, 60, 90];
        return (
          <div> <h4> Book This Room Now: </h4>
            <FFButton ButtonTitle="Book Now" DropdownContent={DropdownContent} BtnFunc="BookNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
          </div>
        );
      }
      else if (timeDifference > 60) {
        let DropdownContent = [15, 30, 45, 60];
        return (
          <div> <h4> Book This Room Now: </h4>
            <FFButton ButtonTitle="Book Now" DropdownContent={DropdownContent} BtnFunc="BookNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
          </div>
        );
      }
      else if (timeDifference > 45) {
        let DropdownContent = [15, 30, 45];
        return (
          <div> <h4> Book This Room Now: </h4>
            <FFButton ButtonTitle="Book Now" DropdownContent={DropdownContent} BtnFunc="BookNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
          </div>
        );
      }
      else if (timeDifference > 30) {
        let DropdownContent = [15, 30];
        return (
          <div> <h4> Book This Room Now: </h4>
            <FFButton ButtonTitle="Book Now" DropdownContent={DropdownContent} BtnFunc="BookNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
          </div>
        );
      }
      else if (timeDifference > 15) {
        let DropdownContent = [15, timeDifference];
        return (
          <div> <h4> Book This Room Now: </h4>
            <FFButton ButtonTitle="Book Now" DropdownContent={DropdownContent} BtnFunc="BookNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
          </div>
        );
      }
      else if (timeDifference >= 5) {
        let DropdownContent = [timeDifference];
        return (
          <div> <h4> Book This Room Now: </h4>
            <FFButton ButtonTitle="Book Now" DropdownContent={DropdownContent} BtnFunc="BookNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
          </div>
        );
      }
      else {
        if (room.Appointments[1]) {
          let currentAppointmentEnd = new Date(parseInt(room.Appointments[0].End, 10));
          now = moment(now);
          currentAppointmentEnd = moment(currentAppointmentEnd);
          let nextAppointmentStart = new Date(parseInt(room.Appointments[1].Start, 10));
          nextAppointmentStart = moment(nextAppointmentStart);
          let DropdownContent;
          let timeDifference = nextAppointmentStart.diff(currentAppointmentEnd, 'minutes')
          if (timeDifference > 120) {
            DropdownContent = [15, 30, 60, 120]
            return (
              <div> <h4> Book after next meeting Now: </h4>
                <FFButton ButtonTitle="Book After Next" DropdownContent={DropdownContent} BtnFunc="BookAfterNext" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
              </div>
            );
          }
          else if (timeDifference > 90) {
            DropdownContent = [15, 30, 60, 90];
            return (
              <div> <h4> Book after next meeting Now: </h4>
                <FFButton ButtonTitle="Book After Next" DropdownContent={DropdownContent} BtnFunc="BookAfterNext" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
              </div>
            );
          }
          else if (timeDifference > 60) {
            DropdownContent = [15, 30, 45, 60];
            return (
              <div> <h4> Book after next meeting Now: </h4>
                <FFButton ButtonTitle="Book After Next" DropdownContent={DropdownContent} BtnFunc="BookAfterNext" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
              </div>
            );
          }
          else if (timeDifference > 30) {
            DropdownContent = [15, 30, timeDifference];
            return (
              <div> <h4> Book after next meeting Now: </h4>
                <FFButton ButtonTitle="Book After Next" DropdownContent={DropdownContent} BtnFunc="BookAfterNext" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
              </div>
            );
          }
          else if (timeDifference > 15) {
            DropdownContent = [15, timeDifference];
            return (
              <div> <h4> Book after next meeting Now: </h4>
                <FFButton ButtonTitle="Book After Next" DropdownContent={DropdownContent} BtnFunc="BookAfterNext" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
              </div>
            );
          }
          else if (timeDifference > 5) {
            DropdownContent = [timeDifference];
            return (
              <div> <h4> Book after next meeting Now: </h4>
                <FFButton ButtonTitle="Book After Next" DropdownContent={DropdownContent} BtnFunc="BookAfterNext" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
              </div>
            );
          }
          else {
            return (<div> <h4> Room is unavailable </h4></div>);
          }
        }
        else {
          let DropdownContent = [15, 30, 60, 120]
          return (
            <div> <h4> Book after next meeting Now: </h4>
              <FFButton ButtonTitle="Book After Next" DropdownContent={DropdownContent} BtnFunc="BookAfterNext" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
            </div>
          );
        }
      }
    }
    else {
      let DropdownContent = [15, 30, 60, 120];
      return (
        <div> <h4> Book This Room Now: </h4>
          <FFButton ButtonTitle="Book Now" DropdownContent={DropdownContent} BtnFunc="BookNow" room={room} togglePopup={props.togglePopup} showPopup={props.showPopup} />
        </div>
      );
    }
  }
}
const RoomStatusBlock = ({ config, details, room, togglePopup, showPopup }) => (
  <div className={room.Busy ? 'columns small-8 left-col busy' : 'columns small-8 left-col open'}>
    <table id="single-room__logo-name-table">
      <tr><td id="single-room__org-logo-td">
        <div id="single-room__org-logo">
          <img src="../img/logo.png" alt="Logo" id="single-room__org-logo__img" />
        </div>
      </td>
        <td>
          <div id="single-room__room-name">{room.Name}</div>
        </td></tr>
    </table>
    <div id="single-room__room-status">{room.Busy ? config.statusBusy : config.statusAvailable}</div>
    <ButtonControl room={room} details={details} togglePopup={togglePopup} showPopup={showPopup} />
    <Details room={room} details={details} />
    <Time room={room} details={details} />
    <Organizer room={room} details={details} />
    {process.env.REACT_APP_FLIGHTBOARD_MODE == "true" &&
      <Link to={'../'} target="_self">
        <li className="chevron">Back to Flightboard</li>
      </Link>
    }

  </div>
);

RoomStatusBlock.propTypes = {
  room: PropTypes.object.isRequired,
  details: PropTypes.object,
  config: PropTypes.object,
  togglePopup: PropTypes.func,
  showPopup: PropTypes.bool
}

export default RoomStatusBlock;

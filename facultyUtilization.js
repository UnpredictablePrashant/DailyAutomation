//App Script Code

function exportCalendarToCSV() {

  var calendar = CalendarApp.getDefaultCalendar();
  var startDate = new Date('2024-10-01T00:00:00');
  var endDate = new Date('2024-10-31T23:59:59');
  var events = calendar.getEvents(startDate, endDate);
  var csvContent = 'Date,Duration (minutes),Participant Name,Participant Email,Meeting Title\n';
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    var title = event.getTitle();
    var startTime = event.getStartTime();
    var endTime = event.getEndTime();
    var duration = (endTime - startTime) / (1000 * 60);
    var organizerEmail = event.getCreators();
    
    if (organizerEmail === 'admissions@herovired.com') {
      title = 'Sales Screening Calls';
    }
    
    var attendees = event.getGuestList();
    var participantName = '';
    var participantEmail = '';
    
    var externalParticipantFound = false;

    for (var j = 0; j < attendees.length; j++) {
      var attendee = attendees[j];
      var email = attendee.getEmail();
      
      if (!email.includes('herovired.com')) {
        participantName = attendee.getName() && attendee.getName() !== "Unknown" ? attendee.getName() : email;
        participantEmail = email;
        externalParticipantFound = true;
        break;  
      }
    }
    
    if (!externalParticipantFound) {
      participantName = event.getCreators() || organizerEmail;
      participantEmail = organizerEmail;
    }
    
    csvContent += [startTime.toISOString().slice(0, 10), duration, participantName, participantEmail, title].join(',') + '\n';
  }
  
  var fileName = 'October_2024_Calendar_Events.csv';
  var file = DriveApp.createFile(fileName, csvContent);
  
  Logger.log('CSV file created: ' + file.getUrl());
}

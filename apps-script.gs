function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = e.parameter;

  var headers = ['Timestamp', 'Full Name', 'Phone', 'Email', 'Location', 'School',
    'Study Level', 'Course', 'Experience Level', 'Programming Languages',
    'Tech Stack', 'Why Join', 'Goals', 'Availability', 'Laptop Access',
    'Emergency Contact', 'Confirmed'];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  sheet.appendRow([
    new Date(),
    data.fullName || '',
    data.phone || '',
    data.email || '',
    data.location || '',
    data.school || '',
    data.studyLevel || '',
    data.course || '',
    data.experienceLevel || '',
    data.languages || '',
    data.techStack || '',
    data.whyJoin || '',
    data.goals || '',
    data.availability || '',
    data.laptop || '',
    data.emergency || '',
    data.commitment || ''
  ]);

  return ContentService.createTextOutput('Success');
}

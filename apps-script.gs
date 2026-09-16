var ADMIN_EMAIL = 'info@tsi.ac.ke';
var LOGO_URL = 'https://ai-bootcamp-delta.vercel.app/tsi.png';
var INSTITUTE_NAME = 'Techsavanna Software Institute';

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

  try {
    sendApplicantConfirmation(data);
  } catch (err) {
    // A mail failure should never block the applicant from seeing success,
    // since their data is already safely recorded above.
  }

  try {
    sendAdminNotification(data, sheet.getParent().getUrl());
  } catch (err) {
    // Same reasoning: don't fail the whole request over a notification email.
  }

  return ContentService.createTextOutput('Success');
}

function sendApplicantConfirmation(data) {
  if (!data.email) return;

  var subject = 'Your application to the Free Software Engineering Bootcamp';
  var html = ''
    + '<div style="font-family:Arial,Helvetica,sans-serif; max-width:560px; margin:0 auto; color:#1a1a1a;">'
    + '<img src="' + LOGO_URL + '" alt="' + INSTITUTE_NAME + '" style="max-width:200px; height:auto; margin-bottom:24px;">'
    + '<h2 style="color:#4f46e5; margin-bottom:4px;">Application Received</h2>'
    + '<p>Hi ' + escapeHtml(data.fullName || 'there') + ',</p>'
    + '<p>Thank you for applying to the Free Software Engineering Bootcamp at ' + INSTITUTE_NAME + '. '
    + 'We have received your application and it is now under review.</p>'
    + '<table style="width:100%; border-collapse:collapse; margin:20px 0; font-size:14px;">'
    + summaryRow('Full Name', data.fullName)
    + summaryRow('Phone', data.phone)
    + summaryRow('Email', data.email)
    + summaryRow('School/University', data.school)
    + summaryRow('Experience Level', data.experienceLevel)
    + '</table>'
    + '<p><strong>What happens next:</strong></p>'
    + '<ul style="padding-left:20px; line-height:1.8;">'
    + '<li>Our team will review your application</li>'
    + '<li>You will be contacted via WhatsApp or phone</li>'
    + '<li>Onboarding details will follow before the bootcamp starts</li>'
    + '</ul>'
    + '<p style="margin-top:24px; color:#555;">If you did not submit this application, you can safely ignore this email.</p>'
    + '<p style="margin-top:32px;">Best regards,<br><strong>' + INSTITUTE_NAME + '</strong></p>'
    + '</div>';

  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    htmlBody: html
  });
}

function sendAdminNotification(data, sheetUrl) {
  var subject = 'New Bootcamp Application: ' + (data.fullName || 'Unknown applicant');
  var html = ''
    + '<div style="font-family:Arial,Helvetica,sans-serif; max-width:560px; margin:0 auto; color:#1a1a1a;">'
    + '<h2 style="color:#4f46e5;">New Enrolment Received</h2>'
    + '<p>A new application was submitted to the Free Software Engineering Bootcamp.</p>'
    + '<table style="width:100%; border-collapse:collapse; margin:20px 0; font-size:14px;">'
    + summaryRow('Full Name', data.fullName)
    + summaryRow('Phone', data.phone)
    + summaryRow('Email', data.email)
    + summaryRow('Location', data.location)
    + summaryRow('School/University', data.school)
    + summaryRow('Study Level', data.studyLevel)
    + summaryRow('Course/Subjects', data.course)
    + summaryRow('Experience Level', data.experienceLevel)
    + summaryRow('Programming Languages', data.languages)
    + summaryRow('Tech Stack', data.techStack)
    + summaryRow('Why Join', data.whyJoin)
    + summaryRow('Goals', data.goals)
    + summaryRow('Availability', data.availability)
    + summaryRow('Laptop Access', data.laptop)
    + summaryRow('Emergency Contact', data.emergency)
    + '</table>'
    + '<p><a href="' + sheetUrl + '">Open the full spreadsheet</a></p>'
    + '</div>';

  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: subject,
    htmlBody: html
  });
}

function summaryRow(label, value) {
  return '<tr>'
    + '<td style="padding:6px 10px 6px 0; color:#666; white-space:nowrap; vertical-align:top;"><strong>' + label + '</strong></td>'
    + '<td style="padding:6px 0; vertical-align:top;">' + escapeHtml(value || '-') + '</td>'
    + '</tr>';
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function doGet(e) {
  return ContentService.createTextOutput('This endpoint only accepts POST requests.');
}

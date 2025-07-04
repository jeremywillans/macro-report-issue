/* eslint-disable no-undef */
/* eslint-disable no-nested-ternary */
/* eslint-disable no-console */
/*
# Report Issue Macro
# Written by Jeremy Willans
# https://github.com/jeremywillans/macro-report-issue
#
# USE AT OWN RISK, MACRO NOT FULLY TESTED NOR SUPPLIED WITH ANY GUARANTEE
#
# Usage -

#  This macro will provide the following options
#  -  Show a survey at the end of each call to capture user experience
#  -  Allow users to report an issue directly from the Touch Panel
#
#  The data can be made available to the following destinations
#  Webex Space, Microsoft Teams Channel, HTTP Server (POST) and/or Service Now Incident.
#
*/
// eslint-disable-next-line import/no-unresolved
import xapi from 'xapi';

const version = '0.0.6';

// Macro Options
const o = {
  appName: 'reportIssue', // Prefix used for Logging and UI Extensions
  widgetPrefix: 'ri-', // Prefix used for UI Widget and Feedback Identifiers
  debugButton: false, // Enables deployment of call survey debug button designed for testing
  // Call Parameters
  callEnabled: false, // Should calls be processed (disable to only use button)
  minDuration: 10, // Minimum call duration (seconds) before Survey is displayed
  // Panel Parameters
  panelEmoticons: true, // Show emoticons on the panel
  panelComments: true, // Show comments on the panel
  panelTips: true, // Show text tips for category and issue selections
  // User Selection
  userParseEmail: false, // Parse and validate an email address in user field
  userSuggest: true, // Enables suggestion to enter reporter when enabling ticket option
  // Button Parameters
  buttonEnabled: true, // Include a Report Issue button on screen
  buttonLocation: 'HomeScreenAndCallControls', // Valid HomeScreen,HomeScreenAndCallControls,ControlPanel
  buttonPosition: 1, // Button order position
  buttonIcon: 'Concierge', // Icon for Report Issue / Debug buttons
  buttonColor: '#1170CF', // Color of button, default blue
  // Webex Space Parameters
  webexEnabled: false, // Enable for Webex Space Message Logging
  webexLogExcellent: false, // Optionally log excellent results to Webex Space
  webexBotToken: '', // Webex Bot Token for sending messages
  webexRoomId: '', // Webex Room Id for sending messages
  webexReportRoomId: '', // If defined, Report Issue messages will be sent here.
  webexWebhook: '', // Webex Incoming Webhook for sending messages
  webexReportWebhook: '', // If defined, Report Issue messages will be sent to this webhook.
  // MS Teams Channel Parameters
  teamsEnabled: false, // Send message to MS Teams channel when room released
  teamsLogExcellent: false, // Optionally log excellent results to Teams Channel
  teamsWebhook: '', // URL for Teams Channel Incoming Webhook
  teamsReportWebhook: '', // If defined, Report Issue messages will be sent here.
  // HTTP JSON Post Parameters
  httpEnabled: false, // Enable for JSON HTTP POST Destination
  httpUrl: '', // HTTP Server POST URL
  httpAuth: false, // Destination requires HTTP Header for Authentication
  httpHeader: 'Authorization: XXXX', // Header Content for HTTP POST Authentication
  // Service Now Parameters
  snowEnabled: false, // Enable for Service NOW Incident Raise
  snowTicketCall: true, // Enabled UI Checkbox to Raise Ticket for Call Survey
  snowTicketReport: false, // Enable UI Checkbox to Raise Ticket for Report Issue
  snowRaiseAverage: false, // Enabled to raise Incident for Average Survey response (3/4 stars)
  snowUserLookup: true, // Lookup user in SNOW when entered (required for snowUserRequired)
  snowUserAppend: '', // Allows appending string (such as @domain) during lookup
  snowUserRequired: false, // Require user when raising a ticket in SNOW
  snowUserField: 'user_name', // Field to perform lookup from Service Now (alt is email)
  //
  // Note: snowRaiseAverage is overridden by snowTicketCall if enabled.
  //
  snowInstance: 'xxxx.service-now.com', // Specify the base url for Service Now
  snowCredentials: '', // Basic Auth format is "username:password" base64-encoded
  snowCallerId: '', // Default Caller for Incidents, needs to be sys_id of Caller
  snowCmdbCi: '', // Default CMDB CI, needs to be sys_id of CI
  snowCmdbLookup: true, // Lookup Device using Serial Number in Service Now
  snowExtra: { // Any extra parameters to pass to Service Now
    // assignment_group: '', // Field value can be Name or sys_id
  },
  // Global Parameters
  defaultSubmit: true, // Send post call results if not explicitly submitted (timeout)
  uploadLogsCallPoor: true, // Enables auto uploading of logs when a poor rating is given,
  uploadLogsCallAverage: false, // Enables auto uploading of logs when an average rating is given
  uploadLogsReport: false, // Enables auto uploading of logs when submitting a report issue
  // Timeout Parameters
  timeoutPanel: 20,
  timeoutPopup: 10,
  // Logging Parameters
  logDetailed: true, // Enable detailed logging
  logUnknownResponses: false, // Show unknown extension responses in the log
};

// Language / Text Options
const l10n = {
  categoryTip: 'Please select the most appropriate category', // Tip shown when selecting a category
  issueTip: 'Please select the most appropriate issue', // Tip shown when selecting an issue
  buttonText: 'Report Issue', // Text of button on Touch Panel
  issuePrefix: 'Report Issue', // Prefix shown for report issue titles
  feedbackPrefix: 'Feedback', // Prefix shown for call survey feedback titles
  ticketTerm: 'Incident', // Button terminology used for the Ticket
  userField: 'Username', // Terminology used on panel to represent user field
  userPlaceholder: 'Please provide your username', // Placeholder for User prompt
  navigatorText: 'Please complete Feedback Survey on the Touchpanel', // Alert message shown on Board if Navigator connected
};

// Maximum of 4 Categories
// Maximum of 5 Issues per Category
const categories = {
  video: {
    text: 'Video Issue',
    prompt: `${o.panelEmoticons ? '📹 ' : ''}Video`,
    issues: [
      { id: 'inbound-video', text: 'Issue with remote video' },
      { id: 'outbound-video', text: 'Remote participants cannot see me' },
      { id: 'video-quality', text: 'Bad video quality' },
      { id: 'other', text: 'Other' },
    ],
    // Category level SNOW field
    // snowExtra: { assignment_group: 'sys_id-of-assignment-group' },
  },
  audio: {
    text: 'Audio Issue',
    prompt: `${o.panelEmoticons ? '🎙️ ' : ''}Audio`,
    issues: [
      { id: 'inbound-audio', text: 'Issue with remote audio' },
      { id: 'outbound-audio', text: 'Remote participants cannot hear me' },
      { id: 'audio-quality', text: 'Bad audio quality' },
      { id: 'other', text: 'Other' },
    ],
    // Category level SNOW field
    // snowExtra: { assignment_group: 'sys_id-of-assignment-group' },
  },
  share: {
    text: 'Share Issue',
    prompt: `${o.panelEmoticons ? '📺 ' : ''}Share`,
    issues: [
      { id: 'share-issue', text: 'Unable to share' },
      { id: 'outbound-content', text: 'Remote participants cannot see my content' },
      { id: 'content-quality', text: 'Bad content quality' },
      { id: 'other', text: 'Other' },
    ],
    // Category level SNOW field
    // snowExtra: { assignment_group: 'sys_id-of-assignment-group' },
  },
  room: {
    text: 'Room Equipment',
    prompt: `${o.panelEmoticons ? '🪑 ' : ''}Room`,
    issues: [
      { id: 'missing-equipment', text: 'Missing cables or equipment' },
      {
        id: 'table-equipment',
        text: 'Broken cables or equipment',
        // Issue level SNOW field
        // snowExtra: { assignment_group: 'sys_id-of-assignment-group' },
      },
      { id: 'cleanliness', text: 'Cleanliness issue' },
      { id: 'environmental', text: 'Environmental issue' },
      { id: 'other', text: 'Other' },
    ],
    // Category level SNOW field
    // snowExtra: { assignment_group: 'sys_id-of-assignment-group' },
  },
};

// Maximum of 3 Ratings
const ratings = [
  // 1 and 2 Stars
  {
    text: 'Poor',
    snowExtra: {
      // urgency: 2,
    },
  },
  // 3 and 4 Stars
  {
    text: 'Average',
    snowExtra: {
      // urgency: 3,
    },
  },
  // 5 Stars,
  { text: 'Excellent' },
];

// ----- EDIT BELOW THIS LINE AT OWN RISK ----- //

const Header = [
  'Content-Type: application/json',
  'Accept: application/json',
];
const webexHeader = [...Header, `Authorization: Bearer ${o.webexBotToken}`];
const snowHeader = [...Header, `Authorization: Basic ${o.snowCredentials}`];
const httpHeader = o.httpAuth ? [...Header, o.httpHeader] : [...Header];
const snowIncidentUrl = `https://${o.snowInstance}/api/now/table/incident`;
const snowUserUrl = `https://${o.snowInstance}/api/now/table/sys_user`;
const snowCMDBUrl = `https://${o.snowInstance}/api/now/table/cmdb_ci`;
const catArray = Object.keys(categories);
const panelId = `${o.appName}-panel-${version.replaceAll('.', '')}`;
const debugSurvey = `${o.appName}-button-debugSurvey-${version.replaceAll('.', '')}`;
const sysInfo = {};
let userInfo = {};
let reportInfo = {};
let callInfo = {};
let callType = '';
let showFeedback = true;
let voluntaryRating = false;
let activeCall = false;
let callMatched = false;
let callDestination = false;
let raiseTicket = false;
let manualReport = true;
let panelTimeout;
let errorResult = false;
let skipLog = true;
let logPending = false;
let panelStage = 0;
let navigatorCheck = false;

// Call Domains
const vimtDomain = '@m.webex.com';
const googleDomain = 'meet.google.com';
const msftDomain = 'teams.microsoft.com';
// eslint-disable-next-line no-useless-escape
const zoomDomain = '(@zm..\.us|@zoomcrc.com)';

// Sleep Function
async function sleep(ms) {
  // eslint-disable-next-line no-promise-executor-return
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Remove UI Elements
async function removePanel(PanelId, showLog = true) {
  if (o.logDetailed && showLog) console.debug(`Removing Panel: ${PanelId}`);
  try {
    await xapi.Command.UserInterface.Extensions.Panel.Remove({ PanelId });
  } catch (error) {
    console.error('Unable to remove Panel');
    console.debug(error.message ? error.message : error);
  }
}

function calculateRating() {
  if (!reportInfo.rating) {
    return 'Overall Rating:     🌟 🌟 🌟 🌟 🌟';
  }
  switch (reportInfo.rating) {
    case 1:
      skipLog = !o.uploadLogsCallPoor;
      return 'Overall Rating:     🌟 🌑 🌑 🌑 🌑';
    case 2:
      skipLog = !o.uploadLogsCallPoor;
      return 'Overall Rating:     🌟 🌟 🌑 🌑 🌑';
    case 3:
      skipLog = !o.uploadLogsCallAverage;
      return 'Overall Rating:     🌟 🌟 🌟 🌑 🌑';
    case 4:
      skipLog = !o.uploadLogsCallAverage;
      return 'Overall Rating:     🌟 🌟 🌟 🌟 🌑';
    default:
      skipLog = true;
      return 'Overall Rating:     🌟 🌟 🌟 🌟 🌟';
  }
}

// Generate UI Panel Dynamically
async function addPanel(newStage = false) {
  if (newStage) panelStage = newStage;
  const prefix = manualReport ? l10n.issuePrefix : l10n.feedbackPrefix;
  let title = `${prefix} - Select a Category`;
  let selectedIssue = false;
  if (panelStage === 2) {
    title = `${prefix} - Additional Fields`;
    selectedIssue = categories[reportInfo.category].issues.find((i) => i.id === reportInfo.issue);
  } else if (panelStage === 1) {
    title = `${prefix} - Select an Issue`;
  }

  const userDisplay = reportInfo.reporter_display || reportInfo.reporter;
  const userStatus = raiseTicket ? o.snowUserRequired ? 'Required' : 'Optional' : 'Optional';

  const xml = `
    <Extensions>
      <Version>1.11</Version>
      <Panel>
        <Order>${o.buttonPosition}</Order>
        <PanelId>${panelId}</PanelId>
        <Location>${o.buttonEnabled ? sysInfo.isRoomOS ? o.buttonLocation : 'ControlPanel' : 'Hidden'}</Location>
        <Icon>Concierge</Icon>
        <Color>${o.buttonColor}</Color>
        <Name>${l10n.buttonText}</Name>
        <ActivityType>Custom</ActivityType>
        <Page>
          <Name>${title}</Name>
          ${manualReport ? '' : `<Row>
            <Name/>
            <Widget>
              <WidgetId>${o.widgetPrefix}rating_text</WidgetId>
              <Name>${calculateRating()}</Name>
              <Type>Text</Type>
              <Options>size=3;fontSize=normal;align=center</Options>
            </Widget>
            <Widget>
              <WidgetId>${o.widgetPrefix}rating_edit</WidgetId>
              <Name>Edit</Name>
              <Type>Button</Type>
              <Options>size=1</Options>
            </Widget>
          </Row>`}
          <Row>
            <Name/>
            ${o.panelTips && panelStage === 0 ? `
            <Widget>
              <WidgetId>${o.widgetPrefix}category_text</WidgetId>
              <Name>${l10n.categoryTip}</Name>
              <Type>Text</Type>
              <Options>size=3;fontSize=normal;align=center</Options>
            </Widget>
            ` : ''}
            <Widget>
              <WidgetId>${o.widgetPrefix}category</WidgetId>
              <Type>GroupButton</Type>
              <Options>size=${catArray.length + 1}${panelStage >= 1 ? '' : catArray.length > 3 ? ';columns=2' : ''}</Options>
              <ValueSpace>
                <Value>
                  <Key>${catArray[0]}</Key>
                  <Name>${categories[catArray[0]].prompt}</Name>
                </Value>
                <Value>
                  <Key>${catArray[1]}</Key>
                  <Name>${categories[catArray[1]].prompt}</Name>
                </Value>
                ${catArray[2] ? `
                <Value>
                  <Key>${catArray[2]}</Key>
                  <Name>${categories[catArray[2]].prompt}</Name>
                </Value>
                ` : ''}
                ${catArray[3] ? `
                <Value>
                  <Key>${catArray[3]}</Key>
                  <Name>${categories[catArray[3]].prompt}</Name>
                </Value>
                ` : ''}
              </ValueSpace>
            </Widget>
          </Row>
          ${panelStage > 0 ? panelStage === 2 ? `
          <Row>
            <Name/>
            <Widget>
              <WidgetId>${o.widgetPrefix}issue-button</WidgetId>
              <Type>Button</Type>
              <Options>size=4</Options>
              <Name>${selectedIssue.text}</Name>
            </Widget>
          </Row>
          ` : `
          <Row>
            <Name/>
            ${o.panelTips && panelStage === 1 ? `
            <Widget>
              <WidgetId>${o.widgetPrefix}category_text</WidgetId>
              <Name>${l10n.issueTip}</Name>
              <Type>Text</Type>
              <Options>size=3;fontSize=normal;align=center</Options>
            </Widget>
            ` : ''}
            <Widget>
              <WidgetId>${o.widgetPrefix}issue</WidgetId>
              <Type>GroupButton</Type>
              <Options>size=4;columns=1</Options>
              <ValueSpace>
                <Value>
                  <Key>${categories[reportInfo.category].issues[0].id}</Key>
                  <Name>${categories[reportInfo.category].issues[0].text}</Name>
                </Value>
                ${categories[reportInfo.category].issues[1] ? `
                <Value>
                  <Key>${categories[reportInfo.category].issues[1].id}</Key>
                  <Name>${categories[reportInfo.category].issues[1].text}</Name>
                </Value>
                ` : ''}
                ${categories[reportInfo.category].issues[2] ? `
                <Value>
                  <Key>${categories[reportInfo.category].issues[2].id}</Key>
                  <Name>${categories[reportInfo.category].issues[2].text}</Name>
                </Value>
                ` : ''}
                ${categories[reportInfo.category].issues[3] ? `
                <Value>
                  <Key>${categories[reportInfo.category].issues[3].id}</Key>
                  <Name>${categories[reportInfo.category].issues[3].text}</Name>
                </Value>
                ` : ''}
                ${categories[reportInfo.category].issues[4] ? `
                <Value>
                  <Key>${categories[reportInfo.category].issues[4].id}</Key>
                  <Name>${categories[reportInfo.category].issues[4].text}</Name>
                </Value>
                ` : ''}
              </ValueSpace>
            </Widget>
          </Row>
          ` : ''}
          ${panelStage === 2 ? `
          <Row>
            <Name/>
            <Widget>
              <WidgetId>${o.widgetPrefix}reporter_text</WidgetId>
              <Name>${o.panelEmoticons ? '👤 ' : ''}${reportInfo.reporter && reportInfo.reporter !== '' ? userDisplay : `${l10n.userField} (${userStatus}) &gt;`}</Name>
              <Type>Text</Type>
              <Options>size=3;fontSize=normal;align=right</Options>
            </Widget>
            <Widget>
              <WidgetId>${o.widgetPrefix}reporter_edit</WidgetId>
              <Name>${reportInfo.reporter && reportInfo.reporter !== '' ? 'Edit' : 'Add'}</Name>
              <Type>Button</Type>
              <Options>size=1</Options>
            </Widget>
          </Row>
          <Row>
            <Name/>
            <Widget>
              <WidgetId>${o.widgetPrefix}comments_text</WidgetId>
              <Name>${o.panelEmoticons ? '💬 ' : ''}${reportInfo.comments && reportInfo.comments !== '' ? 'Edit Comments' : 'Comments (Optional)'} &gt;</Name>
              <Type>Text</Type>
              <Options>size=3;fontSize=normal;align=right</Options>
            </Widget>
            <Widget>
              <WidgetId>${o.widgetPrefix}comments_edit</WidgetId>
              <Name>${reportInfo.comments && reportInfo.comments !== '' ? 'Edit' : 'Add'}</Name>
              <Type>Button</Type>
              <Options>size=1</Options>
            </Widget>
          </Row>
          ${o.panelComments && reportInfo.comments && reportInfo.comments !== '' ? `
          <Row>
            <Name/>
            <Widget>
              <WidgetId>${o.widgetPrefix}comments_result</WidgetId>
              <Name>${reportInfo.comments}</Name>
              <Type>Text</Type>
              <Options>size=4;fontSize=normal;align=center</Options>
            </Widget>
          </Row>
          ` : ''}
          <Row>
            <Name/>
            ${o.snowEnabled && ((!manualReport && o.snowTicketCall) || (manualReport && o.snowTicketReport)) ? `
            <Widget>
              <WidgetId>${o.widgetPrefix}ticket_toggle</WidgetId>
              <Name>${o.panelEmoticons ? '🎟️ ' : ''}Raise ${l10n.ticketTerm}?</Name>
              <Type>Button</Type>
              <Options>size=2</Options>
            </Widget>
            ` : ''}
            <Widget>
              <WidgetId>${o.widgetPrefix}survey_submit</WidgetId>
              <Name>Submit ${o.snowEnabled && raiseTicket ? l10n.ticketTerm : 'Feedback'}${o.panelEmoticons ? ' 🚀' : ''}</Name>
              <Type>Button</Type>
              <Options>size=2</Options>
            </Widget>
          </Row>
          ` : ''}
          <PageId>${panelId}-survey</PageId>
          <Options>hideRowNames=1</Options>
        </Page>
      </Panel>
    </Extensions>
  `;
  await xapi.Command.UserInterface.Extensions.Panel.Save({ PanelId: panelId }, xml);
  if (panelStage === 2) {
    await xapi.Command.UserInterface.Extensions.Widget.SetValue({ WidgetId: `${o.widgetPrefix}issue-button`, Value: 'active' });
  }
  if ((o.snowEnabled && panelStage === 2)
    && ((!manualReport && o.snowTicketCall) || (manualReport && o.snowTicketReport))
  ) {
    await xapi.Command.UserInterface.Extensions.Widget.SetValue({ WidgetId: `${o.widgetPrefix}ticket_toggle`, Value: raiseTicket ? 'active' : 'inactive' });
  }
}

// Reset Variables
function resetVariables() {
  if (o.logDetailed) console.debug('Resetting Panel Variables');
  errorResult = false;
  userInfo = {};
  reportInfo = {};
  showFeedback = true;
  voluntaryRating = false;
  raiseTicket = false;
  manualReport = true;
  skipLog = true;
  panelStage = 0;
  addPanel();
  xapi.Command.UserInterface.Extensions.Widget.UnsetValue({ WidgetId: `${o.widgetPrefix}category` });
  if (!activeCall) {
    console.debug('Resetting Call Variables');
    callMatched = false;
    callDestination = false;
    callInfo = {};
    callType = '';
  } else {
    console.debug('Active Call - Skipping Call Variables Reset');
  }
}

// Category Formatter
function formatCategory(category, type = 'text') {
  if (!categories[category]) return 'Unknown';
  return categories[category][type];
}

// Issue Formatter
function formatIssue(category, issue, type = 'text') {
  if (!categories[category]) return 'Unknown';
  const result = categories[category].issues.find((item) => item.id === issue);
  if (result) {
    return result[type];
  }
  return 'Unknown';
}

// Call Type Formatter
function formatType(type) {
  switch (type) {
    case 'webex':
      return 'Webex';
    case 'endpoint':
      return 'Device/User';
    case 'vimt':
      return 'Teams VIMT';
    case 'msft':
      return 'Teams WebRTC';
    case 'google':
      return 'Google WebRTC';
    case 'zoom':
      return 'Zoom';
    case 'mtr':
      return 'Microsoft Teams Call';
    default:
      return 'Unknown';
  }
}

// Rating Formatter
function formatRating(rating, type = false) {
  const field = type || 'text';
  switch (rating) {
    case 5:
      return ratings[2][field];
    case 4:
    case 3:
      return ratings[1][field];
    case 2:
    case 1:
      return ratings[0][field];
    default:
      return 'Unknown';
  }
}

// Time Formatter
function formatTime(seconds) {
  const d = Math.floor((seconds / 3600) / 24);
  const h = Math.floor((seconds / 3600) % 24);
  const m = Math.floor((seconds % 3600) / 60);
  const s = Math.floor(seconds % 3600 % 60);
  // eslint-disable-next-line no-nested-ternary
  const dDisplay = d > 0 ? d + (d === 1 ? (h > 0 || m > 0 ? ' day, ' : ' day') : (h > 0 || m > 0 ? ' days, ' : ' days')) : '';
  // eslint-disable-next-line no-nested-ternary
  const hDisplay = h > 0 ? h + (h === 1 ? (m > 0 || s > 0 ? ' hour, ' : ' hour') : (m > 0 || s > 0 ? ' hours, ' : ' hours')) : '';
  const mDisplay = m > 0 ? m + (m === 1 ? ' minute' : ' minutes') : '';
  const sDisplay = s > 0 ? s + (s === 1 ? ' second' : ' seconds') : '';

  if (m < 1) {
    return `${dDisplay}${hDisplay}${mDisplay}${sDisplay}`;
  }

  return `${dDisplay}${hDisplay}${mDisplay}`;
}

// Handle Panel Timeout
function setPanelTimeout() {
  clearTimeout(panelTimeout);
  panelTimeout = setTimeout(() => {
    console.debug('Panel Timeout.. closing..');
    xapi.Command.UserInterface.Extensions.Panel.Close();
    resetVariables();
  }, o.timeoutPanel * 1000);
}

// Add Debug Button to Codec
async function addDebugButton() {
  if (o.logDetailed) console.debug(`Adding Debug Button: ${debugSurvey}`);
  const xml = `<?xml version="1.0"?>
  <Extensions>
    <Version>1.11</Version>
    <Panel>
      <PanelId>${debugSurvey}</PanelId>
      <Location>${sysInfo.isRoomOS ? o.buttonLocation : 'ControlPanel'}</Location>
      <Icon>${o.buttonIcon}</Icon>
      <Name>Debug Survey</Name>
      <ActivityType>Custom</ActivityType>
    </Panel>
  </Extensions>`;
  await xapi.Command.UserInterface.Extensions.Panel.Save({ PanelId: debugSurvey }, xml);
}

// Post content to Webex Space
async function postWebex() {
  if (o.logDetailed) console.debug('Process postWebex function');
  let blockquote;
  switch (reportInfo.rating) {
    case 5:
      blockquote = '<blockquote class=success>';
      break;
    case 4:
    case 3:
      blockquote = '<blockquote class=warning>';
      break;
    default:
      blockquote = '<blockquote class=danger>';
  }

  let html = (`<p><b>${manualReport ? l10n.issuePrefix : `${l10n.feedbackPrefix} - ${formatRating(reportInfo.rating)} Report`}</b></p>${blockquote}<b>System Name:</b> ${sysInfo.name}<br><b>Serial Number:</b> ${sysInfo.serial}<br><b>SW Release:</b> ${sysInfo.version}`);
  html += `<br><b>Source:</b> ${manualReport ? 'Report Issue' : 'Call Survey'}`;
  if (callType) { html += `<br><b>Call Type:</b> ${formatType(callType)}`; }
  if (callDestination) { html += `<br><b>Destination:</b> ${callDestination}`; }
  if (callInfo.Duration) { html += `<br><b>Call Duration:</b> ${formatTime(callInfo.Duration)}`; }
  if (callInfo.CauseType) { html += `<br><b>Disconnect Cause:</b> ${callInfo.CauseType}`; }
  if (reportInfo.rating) { html += `<br><b>Rating:</b> ${formatRating(reportInfo.rating)} (${reportInfo.rating})`; }
  if (reportInfo.category) { html += `<br><b>Category:</b> ${formatCategory(reportInfo.category)}`; }
  if (reportInfo.issue) { html += `<br><b>Issue:</b> ${formatIssue(reportInfo.category, reportInfo.issue)}`; }
  if (reportInfo.comments) { html += `<br><b>Comments:</b> ${reportInfo.comments}`; }
  if (!manualReport && o.defaultSubmit) { html += `<br><b>Voluntary Rating:</b> ${voluntaryRating ? 'Yes' : 'No'}`; }
  if (reportInfo.incident) { html += `<br><b>Incident Reference:</b> ${reportInfo.incident}`; }
  if (userInfo.sys_id) {
    html += `<br><b>Reporter:</b>  <a href=webexteams://im?email=${userInfo.email}>${userInfo.name}</a> (${userInfo.email})`;
  } else if (reportInfo.reporter) {
    // Include Provided Report Detail if not matched in SNOW
    html += `<br><b>Provided ${l10n.userField} :</b> ${reportInfo.reporter}`;
  }
  html += '</blockquote>';

  let roomId = o.webexRoomId;
  if (manualReport && (o.webexReportRoomId && o.webexReportRoomId !== '')) {
    roomId = o.webexReportRoomId;
  }

  if (roomId && roomId !== '') {
    if (o.logDetailed) console.debug('Sending Webex using Bot Token');
    const messageContent = { roomId, html };

    try {
      const result = await xapi.Command.HttpClient.Post(
        { Header: webexHeader, Url: 'https://webexapis.com/v1/messages' },
        JSON.stringify(messageContent),
      );
      if (/20[04]/.test(result.StatusCode)) {
        if (o.logDetailed) console.debug('postWebex message sent.');
        return;
      }
      console.error(`postWebex status: ${result.StatusCode}`);
      errorResult = true;
      if (result.message && o.logDetailed) {
        console.debug(`${result.message}`);
      }
    } catch (error) {
      console.error('postWebex error');
      console.debug(error.message ? error.message : error);
      errorResult = true;
    }
  }

  let webhook = o.webexWebhook;
  if (manualReport && (o.webexReportWebhook && o.webexReportWebhook !== '')) {
    webhook = o.webexReportWebhook;
  }

  if (webhook && webhook !== '') {
    if (o.logDetailed) console.debug('Sending Webex using Webhook');
    const messageContent = { markdown: html };
    try {
      const result = await xapi.Command.HttpClient.Post(
        { Header, Url: webhook },
        JSON.stringify(messageContent),
      );
      if (/20[04]/.test(result.StatusCode)) {
        if (o.logDetailed) console.debug('postWebex message sent.');
        return;
      }
      console.error(`postWebex status: ${result.StatusCode}`);
      errorResult = true;
      if (result.message && o.logDetailed) {
        console.debug(`${result.message}`);
      }
    } catch (error) {
      console.error('postWebex error');
      console.debug(error.message ? error.message : error);
      errorResult = true;
    }
  }
}

// Post content to MS Teams Channel
async function postTeams() {
  if (o.logDetailed) console.debug('Process postTeams function');
  let color;
  switch (reportInfo.rating) {
    case 5:
      color = 'Good';
      break;
    case 4:
    case 3:
      color = 'Warning';
      break;
    default:
      color = 'Attention';
  }

  const cardBody = {
    type: 'message',
    attachments: [
      {
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: {
          $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
          type: 'AdaptiveCard',
          version: '1.3',
          body: [
            {
              type: 'TextBlock',
              text: manualReport ? l10n.issuePrefix : `${l10n.feedbackPrefix} - ${formatRating(reportInfo.rating)} Report`,
              weight: 'Bolder',
              size: 'Medium',
              color,
            },
            {
              type: 'FactSet',
            },
          ],
        },

      },
    ],
  };

  const facts = [
    {
      title: 'System Name',
      value: sysInfo.name,
    },
    {
      title: 'Serial Number',
      value: sysInfo.serial,
    },
    {
      title: 'SW Release',
      value: sysInfo.version,
    },
    {
      title: 'Source',
      value: manualReport ? 'Report Issue' : 'Call Survey',
    },
  ];

  if (callType) facts.push({ title: 'Call Type', value: formatType(callType) });
  if (callDestination) facts.push({ title: 'Destination', value: callDestination });
  if (callInfo.Duration) facts.push({ title: 'Call Duration', value: formatTime(callInfo.Duration) });
  if (callInfo.CauseType) facts.push({ title: 'Disconnect Cause', value: callInfo.CauseType });
  if (reportInfo.rating) facts.push({ title: 'Rating:', value: `${formatRating(reportInfo.rating)} (${reportInfo.rating})` });
  if (reportInfo.category) facts.push({ title: 'Category', value: formatCategory(reportInfo.category) });
  if (reportInfo.issue) facts.push({ title: 'Issue', value: formatIssue(reportInfo.category, reportInfo.issue) });
  if (!manualReport && o.defaultSubmit) facts.push({ title: 'Voluntary Rating', value: voluntaryRating ? 'Yes' : 'No' });
  if (reportInfo.incident) facts.push({ title: 'Incident Reference', value: reportInfo.incident });

  if (userInfo.sys_id) {
    facts.push({ title: 'Reporter', value: `${userInfo.name} (${userInfo.email})` });
  } else if (reportInfo.reporter) {
    // Include Provided Email if not matched in SNOW
    facts.push({ title: `Provided ${l10n.userField}`, value: reportInfo.reporter });
  }

  cardBody.attachments[0].content.body[1].facts = facts;

  if (reportInfo.comments) {
    cardBody.attachments[0].content.body.push({ type: 'TextBlock', text: 'Comments', weight: 'Bolder' });
    cardBody.attachments[0].content.body.push({ type: 'TextBlock', text: reportInfo.comments, wrap: true });
  }

  let webhook = o.teamsWebhook;
  if (manualReport && (o.teamsReportWebhook && o.teamsReportWebhook !== '')) {
    webhook = o.teamsReportWebhook;
  }

  try {
    const result = await xapi.Command.HttpClient.Post(
      { Header, Url: webhook },
      JSON.stringify(cardBody),
    );
    if (/20[024]/.test(result.StatusCode)) {
      if (o.logDetailed) console.debug('postTeams message sent.');
      return;
    }
    console.error(`postTeams status: ${result.StatusCode}`);
    if (result.message && o.logDetailed) {
      console.debug(`${result.message}`);
    }
  } catch (error) {
    console.error('postTeams error');
    console.debug(error.message ? error.message : error);
  }
}

// Post JSON content to Http Server
async function postHttp() {
  console.debug('Process postHttp function');
  let messageContent = {
    timestamp: Date.now(),
    system: sysInfo.name,
    serial: sysInfo.serial,
    version: sysInfo.version,
    source: manualReport ? 'feedback' : 'call',
    rating: reportInfo.rating ? reportInfo.rating : 0,
    rating_fmt: reportInfo.rating ? formatRating(reportInfo.rating) : '',
    destination: callDestination && callDestination !== '' ? callDestination : '',
    type: callType && callType !== '' ? callType : '',
    type_fmt: callType && callType !== '' ? formatType(callType) : '',
    duration: callInfo.Duration ? callInfo.Duration : 0,
    duration_fmt: callInfo.Duration ? formatTime(callInfo.Duration) : 0,
    cause: callInfo.CauseType ? callInfo.CauseType : '',
    category: reportInfo.category ? reportInfo.category : '',
    category_fmt: reportInfo.category ? formatCategory(reportInfo.category) : '',
    issue: reportInfo.issue ? reportInfo.issue : '',
    issue_fmt: reportInfo.issue ? formatIssue(reportInfo.category, reportInfo.issue) : '',
    comments: reportInfo.comments,
    reporter: reportInfo.reporter,
    voluntary: voluntaryRating ? 1 : 0,
  };

  switch (o.httpFormat) {
    case 'loki':
      messageContent = {
        streams: [
          {
            stream: {
              app: o.appName,
            },
            values: [[`${messageContent.timestamp}000000`, messageContent]],
          },
        ],
      };
      // Append Loki API path if missing.
      if (!o.httpUrl.match('/loki/api/v1/push')) {
        o.httpUrl = o.httpUrl.replace(/\/$/, '');
        o.httpUrl = `${o.httpUrl}/loki/api/v1/push`;
      }
      break;
    case 'powerBi': {
      const ts = new Date(messageContent.timestamp);
      messageContent.timestamp = ts.toISOString();
      messageContent = [messageContent];
      break;
    }
    default:
  }

  try {
    const result = await xapi.Command.HttpClient.Post(
      { Header: httpHeader, Url: o.httpUrl },
      JSON.stringify(messageContent),
    );
    if (/20[04]/.test(result.StatusCode)) {
      if (o.logDetailed) console.debug('postHttp message sent.');
      return;
    }
    console.error(`postHttp status: ${result.StatusCode}`);
    if (result.message && o.logDetailed) {
      console.debug(result.message);
    }
  } catch (error) {
    console.error('postHttp error encountered');
    console.debug(JSON.stringify(error));
  }
}

// Post Incident to Service Now
async function postIncident() {
  if (o.logDetailed) console.debug('Process postIncident function');
  let description = `${manualReport ? l10n.issuePrefix : `${l10n.feedbackPrefix} - ${formatRating(reportInfo.rating)} Report`}\n\nSystem Name: ${sysInfo.name}\nSerial Number: ${sysInfo.serial}\nVersion: ${sysInfo.version}`;
  description += `\nSource: ${manualReport ? 'Report Issue' : 'Call Survey'}`;
  if (callType) { description += `\nCall Type: ${formatType(callType)}`; }
  if (callDestination) { description += `\nDestination: \`${callDestination}\``; }
  if (callInfo.Duration) { description += `\nCall Duration: ${formatTime(callInfo.Duration)}`; }
  if (callInfo.CauseType) { description += `\nDisconnect Cause: ${callInfo.CauseType}`; }
  if (reportInfo.rating) { description += `\n\nRating: ${formatRating(reportInfo.rating)} (${reportInfo.rating})`; }
  if (reportInfo.category) { description += `\nCategory: ${formatCategory(reportInfo.category)}`; }
  if (reportInfo.issue) { description += `\nIssue: ${formatIssue(reportInfo.category, reportInfo.issue)}`; }
  if (reportInfo.comments) { description += `\nComments: ${reportInfo.comments}`; }
  const shortDescription = `${sysInfo.name}: ${manualReport ? l10n.issuePrefix : l10n.feedbackPrefix}`;

  // Initial Construct Incident
  let messageContent = { short_description: shortDescription, description };
  // Add Default Caller, if defined.
  if (o.snowCallerId) {
    messageContent.caller_id = o.snowCallerId;
  }

  if (userInfo.sys_id) {
    messageContent.caller_id = userInfo.sys_id;
  } else if (reportInfo.reporter) {
    messageContent.description += `\nProvided ${l10n.userField}: ${reportInfo.reporter}}`;
  }

  if (o.snowCmdbCi) {
    messageContent.cmdb_ci = o.snowCmdbCi;
  }

  if (o.snowCmdbLookup && sysInfo.cmdbCi) {
    messageContent.cmdb_ci = sysInfo.cmdbCi;
  }

  // Merge Extra Params from Default Options
  if (o.snowExtra) {
    messageContent = { ...messageContent, ...o.snowExtra };
  }

  const q = reportInfo;

  // Merge Extra Params from Selected Category
  if (q.category && categories[q.category].snowExtra) {
    messageContent = { ...messageContent, ...categories[q.category].snowExtra };
  }

  // Merge Extra Params from Selected Category Issue
  if (q.category) {
    const findIssue = categories[q.category].issues.find((item) => item.id === q.issue);
    if (findIssue && findIssue.snowExtra) {
      messageContent = { ...messageContent, ...findIssue.snowExtra };
    }
  }

  // Merge Extra Params from Selected Rating
  const ratingSnowExtra = formatRating(q.rating, 'snowExtra');
  if (ratingSnowExtra) {
    messageContent = { ...messageContent, ...ratingSnowExtra };
  }

  try {
    if (o.logDetailed) console.debug(JSON.stringify(messageContent));
    let result = await xapi.Command.HttpClient.Post(
      { Header: snowHeader, Url: snowIncidentUrl },
      JSON.stringify(messageContent),
    );
    const incidentUrl = result.Headers.find((x) => x.Key === 'Location').Value;
    result = await xapi.Command.HttpClient.Get({ Header: snowHeader, Url: incidentUrl });
    reportInfo.incident = JSON.parse(result.Body).result.number;
    if (o.logDetailed) console.debug(`postIncident successful: ${reportInfo.incident}`);
  } catch (error) {
    console.error('postIncident error encountered');
    console.debug(error.message ? error.message : error);
    errorResult = true;
  }
}

// Close panel and Process enabled services
async function processRequest() {
  if (o.logDetailed) console.debug('Processing Request');
  clearTimeout(panelTimeout);
  await xapi.Command.UserInterface.Extensions.Panel.Close();
  if (o.httpEnabled) {
    postHttp();
  }
  if (o.snowEnabled && raiseTicket) {
    await postIncident();
  }
  if (o.webexEnabled && (
    // Always post manual reports
    manualReport
    // Post if rating is Excellent and logging is enabled
    || ((reportInfo.rating === 1 && o.webexLogExcellent)
    // Post if rating is Average or Poor Rating
    || reportInfo.rating !== 1
    // Always post if contains Comments
    || (reportInfo.comments && reportInfo.comments !== '')))
  ) {
    await postWebex();
  }
  if (o.teamsEnabled && (
    // Always post manual reports
    manualReport
    // Post if rating is Excellent and logging is enabled
    || ((reportInfo.rating === 1 && o.teamsLogExcellent)
    // Post if rating is Average or Poor Rating
    || reportInfo.rating !== 1
    // Always post if contains Comments
    || (reportInfo.comments && reportInfo.comments !== '')))
  ) {
    await postTeams();
  }
  // Upload Logs if Marked
  if (!skipLog && !logPending) {
    console.debug('Commence Log Upload.');
    logPending = true;
    xapi.Command.Logging.SendLogs()
      .then((result) => {
        console.debug(`Log upload complete${result.LogId ? ` - LogId: ${result.LogId}` : ''}`);
        logPending = false;
      })
      .catch((error) => {
        console.error('Issue uploading log');
        console.debug(error.message ? error.message : error);
        logPending = false;
      });
  }
  await sleep(600);
  if (showFeedback) {
    let Title = 'Acknowledgement';
    let Text = 'Thanks for your feedback!';
    let Duration = 15;
    if (errorResult) {
      Title = 'Error Encountered';
      Text = 'Sorry we were unable to complete this request.<br>Please advise your IT Support team of this error.';
      Duration = 20;
    }
    if (reportInfo.incident) {
      Text += `<br>${l10n.ticketTerm} ${reportInfo.incident} raised.`;
    }
    xapi.Command.UserInterface.Message.Alert.Display({
      Title,
      Text,
      Duration,
    });
  }
  resetVariables();
}

// Process call data
async function processCall() {
  activeCall = true;
  if (callMatched) {
    return;
  }
  let call;
  try {
    [call] = await xapi.Status.Call.get();
  } catch (error) {
    // No Active Call
    return;
  }
  if (!call) return;

  if (call.Protocol && call.Protocol === 'WebRTC') {
    callType = 'webrtc';
    callDestination = call.CallbackNumber;
    // Matched WebRTC Call
    if (call.CallbackNumber.match(msftDomain)) {
      // Matched Teams Call
      callType = 'msft';
      callMatched = true;
      if (o.logDetailed) console.debug(`[${callType}] ${callDestination}`);
      return;
    }
    if (call.CallbackNumber.match(googleDomain)) {
      // Matched Google Call
      callType = 'google';
      callMatched = true;
      if (o.logDetailed) console.debug(`[${callType}] ${callDestination}`);
      return;
    }
    // Fallback WebRTC Call
    if (o.logDetailed) console.debug(`[${callType}] ${callDestination}`);
    return;
  }

  // Default Call Type
  callType = 'sip';
  callDestination = call.CallbackNumber;
  if (call.CallbackNumber.match(vimtDomain)) {
    // Matched VIMT Call
    callType = 'vimt';
    callMatched = true;
    if (o.logDetailed) console.debug(`[${callType}] ${callDestination}`);
    return;
  }
  if (call.CallbackNumber.match('.webex.com')) {
    // Matched Webex Call
    callType = 'webex';
    callMatched = true;
    if (o.logDetailed) console.debug(`[${callType}] ${callDestination}`);
    return;
  }
  if (call.CallbackNumber.match(zoomDomain)) {
    // Matched Zoom Call
    callType = 'zoom';
    callMatched = true;
    if (o.logDetailed) console.debug(`[${callType}] ${callDestination}`);
    return;
  }
  if (call.DeviceType === 'Endpoint' && call.CallbackNumber.match('^[^.]*$')) {
    // Matched Endpoint/User Call
    callType = 'endpoint';
    callDestination = `${call.DisplayName}: ${call.CallbackNumber})`;
    if (o.logDetailed) console.debug(`[${callType}] ${callDestination}`);
    return;
  }
  // Fallback SIP Call
  if (o.logDetailed) console.debug(`[${callType}] ${callDestination}`);
}

// Show Rating Prompt
function showRating(updateRating = false) {
  if (updateRating) clearTimeout(panelTimeout);
  const Text = updateRating ? 'Please select a new rating' : 'How was your call?';
  const Title = 'Room Experience Feedback';
  xapi.Command.UserInterface.Message.Rating.Display({
    Duration: 20, FeedbackId: updateRating ? `${o.widgetPrefix}rating_update` : `${o.widgetPrefix}rating_submit`, Text, Title,
  });
}

// Process after Call Disconnect
function processDisconnect() {
  activeCall = false;
  if (callInfo.Duration > o.minDuration) {
    manualReport = false;
    showRating();
  } else {
    resetVariables();
    /*
    xapi.Command.UserInterface.Message.Prompt.Display({
      Title: l10n.feedbackPrefix,
      Text: 'Call did not complete. What happened?',
      FeedbackId: 'no_call_rating',
      'Option.1': 'I dialled the wrong number!',
      'Option.2': 'Call did not answer',
      'Option.3': 'Oops, wrong button',
    });
    */
  }
}

// Show Text Input to User
function showTextInput(promptId, overrideTitle = false, overrideText = false) {
  // Prevent Survey from closing when prompt open
  clearTimeout(panelTimeout);
  const promptBody = {
    Duration: o.timeoutPopup,
    InputType: 'SingleLine',
    KeyboardState: 'Open',
    SubmitText: 'Submit',
  };
  switch (promptId) {
    case `${o.widgetPrefix}comments_edit`: {
      promptBody.FeedbackId = `${o.widgetPrefix}comments_submit`;
      promptBody.Placeholder = 'Additional Comments';
      promptBody.Text = 'Please provide any additional details';
      promptBody.Title = `${manualReport ? l10n.issuePrefix : l10n.feedbackPrefix} Comments`;
      // Populate Comments if previously added
      if (reportInfo.comments !== '') {
        promptBody.InputText = reportInfo.comments;
      }
      break;
    }
    case `${o.widgetPrefix}reporter_edit`: {
      promptBody.FeedbackId = `${o.widgetPrefix}reporter_submit`;
      promptBody.Placeholder = `Enter your ${l10n.userPlaceholder}`;
      promptBody.Text = `Please provide your ${l10n.userField}`;
      promptBody.Title = `${manualReport ? l10n.issuePrefix : l10n.feedbackPrefix} ${l10n.userField}`;
      // Populate field if previously added
      if (reportInfo.reporter !== '') {
        promptBody.InputText = reportInfo.reporter;
      }
      break;
    }
    default:
      return;
  }
  if (overrideTitle) {
    promptBody.Title = overrideTitle;
  }
  if (overrideText) {
    promptBody.Text = overrideText;
  }
  xapi.Command.UserInterface.Message.TextInput.Display(promptBody);
}

// Show Alert Prompt to User
function showAlert(promptId, overrideTitle = false, overrideText = false) {
  // Prevent Survey from closing when prompt open
  clearTimeout(panelTimeout);
  const promptBody = {
    Duration: o.timeoutPopup,
  };
  switch (promptId) {
    case `${o.widgetPrefix}reporter_missing`: {
      promptBody.Text = `Please provide a valid ${l10n.userField}.`;
      promptBody.Title = '⚠️ Missing Reporter ⚠️';
      break;
    }
    case `${o.widgetPrefix}reporter_invalid`: {
      promptBody.Text = `Unable to match the provided ${l10n.userField}<br>Please try again.`;
      promptBody.Title = '⚠️ Invalid Reporter ⚠️';
      break;
    }
    default:
      return;
  }
  if (overrideTitle) {
    promptBody.Title = overrideTitle;
  }
  if (overrideText) {
    promptBody.Text = overrideText;
  }
  xapi.Command.UserInterface.Message.Alert.Display(promptBody);
}

// Validate Categories, Issues and Ratings
function processInputs() {
  let result = true;
  if (catArray.length === 1) {
    console.error('Not Enough Categories');
    result = false;
  }
  if (catArray.length > 4) {
    console.error('Too Many Categories');
    result = false;
  }
  catArray.forEach((item) => {
    if (!categories[item].issues || categories[item].issues.length === 0) {
      console.error(`Missing Issues for Category: ${item}`);
      result = false;
    }
    if (categories[item].issues.length > 5) {
      console.error(`Too Many Issues for Category: ${item}`);
      result = false;
    }
  });
  if (Object.keys(ratings).length !== 3) {
    console.error('Must contain 3 Ratings');
    result = false;
  }
  return result;
}

// Process Call Disconnect
xapi.Event.CallDisconnect.on((event) => {
  if (!o.callEnabled) return;
  callInfo = event;
  callInfo.Duration = Number(event.Duration);
  processDisconnect();
});

// Process Outgoing Call Indication
xapi.Event.OutgoingCallIndication.on(() => {
  if (!o.callEnabled) return;
  processCall();
});

// Process Panel Click Events
xapi.Event.UserInterface.Extensions.Widget.Action.on(async (event) => {
  // Skip Clicked and Pressed events
  if (event.Type !== 'released' && event.Type !== 'changed') { return; }
  // Skip Changed event for Sliders
  if (event.Type === 'changed' && !event.WidgetId.includes('toggle')) { return; }
  // Skip non-panel Widgets
  switch (event.WidgetId) {
    case `${o.widgetPrefix}survey_submit`:
      if (raiseTicket && o.snowUserRequired && (!reportInfo.reporter || reportInfo.reporter === '')) {
        setPanelTimeout();
        showAlert(`${o.widgetPrefix}reporter_missing`);
        return;
      }
      if (raiseTicket && o.snowUserRequired && !userInfo.sys_id) {
        setPanelTimeout();
        showAlert(`${o.widgetPrefix}reporter_invalid`);
        return;
      }
      voluntaryRating = true;
      processRequest();
      break;
    case `${o.widgetPrefix}category`:
      setPanelTimeout();
      xapi.Command.UserInterface.Message.TextLine.Clear();
      reportInfo.category = event.Value;
      if (o.logDetailed) console.debug(`Category updated to: ${reportInfo.category}`);
      await addPanel(1);
      await xapi.Command.UserInterface.Extensions.Widget.UnsetValue({ WidgetId: `${o.widgetPrefix}issue` });
      break;
    case `${o.widgetPrefix}issue`:
      setPanelTimeout();
      reportInfo.issue = event.Value;
      if (o.logDetailed) console.debug(`Issue updated to: ${reportInfo.issue}`);
      await addPanel(2);
      if (reportInfo.issue === 'other') showTextInput(`${o.widgetPrefix}comments_edit`);
      break;
    case `${o.widgetPrefix}issue-button`:
      setPanelTimeout();
      await addPanel(1);
      await xapi.Command.UserInterface.Extensions.Widget.UnsetValue({ WidgetId: `${o.widgetPrefix}issue` });
      break;
    case `${o.widgetPrefix}rating_edit`: {
      xapi.Command.UserInterface.Message.TextLine.Clear();
      showRating(true);
      break;
    }
    case `${o.widgetPrefix}ticket_toggle`: {
      setPanelTimeout();
      raiseTicket = !raiseTicket;
      addPanel();
      if (raiseTicket && o.userSuggest && (!reportInfo.reporter || reportInfo.reporter === '')) {
        setPanelTimeout();
        showTextInput(`${o.widgetPrefix}reporter_edit`, '⚠️ Missing Reporter ⚠️', `Please provide your ${l10n.userField} to include in the ${l10n.ticketTerm}`);
        return;
      }
      if (o.logDetailed) console.debug(`Raise Ticket: ${raiseTicket}`);
      break;
    }
    case `${o.widgetPrefix}comments_edit`:
    case `${o.widgetPrefix}reporter_edit`: {
      showTextInput(event.WidgetId);
      break;
    }
    default:
  }
});

// Process Button Click Events
xapi.Event.UserInterface.Extensions.Panel.Clicked.on(async (event) => {
  if (event.PanelId === panelId) {
    raiseTicket = !o.snowTicketReport;
    skipLog = !o.uploadLogsReport;
    setPanelTimeout();
    return;
  }
  if (event.PanelId !== debugSurvey) return;
  if (!o.debugButton) return;
  console.info(`Debug Button Triggered - ${event.PanelId}`);
  callType = sysInfo.isRoomOS ? 'webex' : 'mtr';
  callInfo.Duration = 17;
  if (sysInfo.isRoomOS) {
    callInfo.CauseType = 'LocalDisconnect';
    callDestination = 'spark:123456789@webex.com';
  }
  processDisconnect();
});

// Process Page Closed
xapi.Event.UserInterface.Extensions.Event.PageClosed.on(async (event) => {
  // ignore other page events
  if (event.PageId !== `${panelId}-survey`) return;
  // ignore if survey was submitted
  if (voluntaryRating) return;
  // Process request if default enabled
  if (!manualReport && o.defaultSubmit) {
    showFeedback = false;
    processRequest();
    return;
  }
  await sleep(200);
  clearTimeout(panelTimeout);
  resetVariables();
});

// Process TextInput Response
xapi.Event.UserInterface.Message.TextInput.Response.on(async (event) => {
  switch (event.FeedbackId) {
    case `${o.widgetPrefix}comments_submit`:
      if (event.Text === reportInfo.comments) return;
      reportInfo.comments = event.Text;
      addPanel(2);
      setPanelTimeout();
      break;
    case `${o.widgetPrefix}reporter_submit`:
      if (o.userParseEmail && event.Text !== '' && !/^.*@.*\..*$/.test(event.Text)) {
        await sleep(500);
        console.warn('Invalid Email Address, re-prompting user...');
        await xapi.Command.Audio.Sound.Play({ Sound: 'Binding' });
        showTextInput(`${o.widgetPrefix}reporter_edit`, '⚠️ Invalid Email Address ⚠️');
        return;
      }
      userInfo = {};
      reportInfo.reporter = event.Text;
      reportInfo.reporter_display = false;
      if (o.snowEnabled && (o.snowUserLookup || o.snowUserRequired)) {
        try {
          const user = `${reportInfo.reporter}${o.snowUserAppend}`;
          let result = await xapi.Command.HttpClient.Get(
            { Header: snowHeader, Url: `${snowUserUrl}?sysparm_limit=1&${o.snowUserField}=${user}` },
          );
          result = JSON.parse(result.Body).result;
          // Validate User Data
          switch (result.length) {
            case 0:
              reportInfo.reporter_display = `${reportInfo.reporter} (Not Found)`;
              if (o.logDetailed) console.debug(`SNOW User Not Found - ${user}`);
              break;
            case 1:
              [userInfo] = result;
              reportInfo.reporter_display = `${userInfo.name} (${reportInfo.reporter})`;
              if (o.logDetailed) console.debug(`SNOW User Found - ${userInfo.name}`);
              break;
            default:
              reportInfo.reporter_display = `${reportInfo.reporter} (Multiple Matches)`;
              if (o.logDetailed) console.debug(`Query ${user} matched ${result.length} results`);
          }
        } catch (error) {
          console.error('unable to lookup SNOW user');
          console.debug(error.message ? error.message : error);
        }
      }
      addPanel(2);
      setPanelTimeout();
      break;
    default:
      if (o.logUnknownResponses) console.warn(`Unexpected TextInput.Response: ${event.FeedbackId}`);
  }
});

// Process TextInput Clear
xapi.Event.UserInterface.Message.TextInput.Clear.on((event) => {
  if (event.FeedbackId === '') return;
  switch (event.FeedbackId) {
    case `${o.widgetPrefix}comments_submit`:
    case `${o.widgetPrefix}reporter_submit`:
      setPanelTimeout();
      break;
    default:
      if (o.logUnknownResponses) console.warn(`Unexpected TextInput.Clear: ${event.FeedbackId}`);
  }
});

// Process Rating Response
xapi.Event.UserInterface.Message.Rating.Response.on(async (event) => {
  switch (event.FeedbackId) {
    case `${o.widgetPrefix}rating_submit`:
    case `${o.widgetPrefix}rating_update`:
      if (Number.isNaN(event.Rating)) return;
      reportInfo.rating = Number(event.Rating);
      if (reportInfo.rating === 5) {
        voluntaryRating = true;
        processRequest();
        return;
      }
      addPanel();
      // Determine Raise Ticket status from rating if snowTicketCall disabled
      if (!o.snowTicketCall) {
        // Raise Ticket for Rating 1 and 2
        raiseTicket = true;
        if (reportInfo.rating > 2) {
          // Raise Ticket for Rating 3 and 4 if snowRaiseAverage enabled
          raiseTicket = o.snowRaiseAverage;
        }
      }
      xapi.Command.UserInterface.Message.Rating.Clear({ FeedbackId: event.FeedbackId });
      if (event.FeedbackId === `${o.widgetPrefix}rating_submit`) {
        if (navigatorCheck) {
          let navigators = await xapi.status.get('Peripherals ConnectedDevice');
          navigators = navigators.filter((i) => i.Status === 'Connected');
          const navController = navigators.some((i) => i.Role === 'Controller');
          if (navController) {
            xapi.Command.UserInterface.Message.TextLine.Display({
              Text: l10n.navigatorText,
              Duration: 10,
            });
          }
        }
        xapi.Command.UserInterface.Extensions.Panel.Open({ PanelId: panelId });
        setPanelTimeout();
        return;
      }
      addPanel();
      setPanelTimeout();
      break;
    default:
      if (o.logUnknownResponses) console.warn(`Unexpected Rating.Response: ${event.FeedbackId}`);
  }
});
// Process Rating Clear
xapi.Event.UserInterface.Message.Rating.Cleared.on((event) => {
  if (event.FeedbackId === '') return;
  switch (event.FeedbackId) {
    case `${o.widgetPrefix}rating_submit`:
      // During Rating Response we clear the Rating Prompt, as it lingers in background.
      // If we have already made a selection, don't treat as a default submission, simply return.
      if (reportInfo.rating) return;
      if (!manualReport && o.defaultSubmit) {
        // Process Default Submit as Excellent
        reportInfo.rating = 5;
        showFeedback = false;
        processRequest();
      }
      break;
    case `${o.widgetPrefix}rating_update`:
      setPanelTimeout();
      break;
    default:
      if (o.logUnknownResponses) console.warn(`Unexpected Rating.Clear: ${event.FeedbackId}`);
  }
});

// Process Active Call
xapi.Status.SystemUnit.State.NumberOfActiveCalls.on((status) => {
  if (!o.callEnabled) return;
  let result = status;
  if (result && !Number.isNaN(result)) {
    result = Number(result);
  }
  if (result > 0) {
    processCall();
  }
});

// Process MTR Active Call
xapi.Status.MicrosoftTeams.Calling.InCall.on((status) => {
  if (!o.callEnabled) return;
  const result = /^true$/i.test(status);
  if (result) {
    callType = 'mtr';
    callInfo.startTime = Date.now();
    activeCall = true;
    callMatched = true;
    if (o.logDetailed) console.debug(`[${callType}] Microsoft Teams Call`);
  } else {
    if (callInfo.startTime) {
      try {
        callInfo.Duration = Math.floor(Number(Date.now() - callInfo.startTime) / 1000);
        if (o.logDetailed) console.debug(`MTR Call calculated duration ${callInfo.Duration}s`);
      } catch (error) {
        console.debug('Error calculating MTR Call Duration');
      }
    }
    processDisconnect();
  }
});

function macroReset() {
  // Close any lingering dialogs and remove UI extensions
  xapi.Command.UserInterface.Extensions.Panel.Close();
  xapi.Command.UserInterface.Message.TextInput.Clear();
  removePanel(panelId, false);
  removePanel(debugSurvey, false);
}

// Process Macro Save
xapi.Event.Macros.Macro.Saved.on((event) => {
  if (event.Name === _main_macro_name()) {
    console.info('Reset Panel and Button for Macro Reload');
    macroReset();
  }
});

// Process Macro Deactivated
xapi.Event.Macros.Macro.Deactivated.on((event) => {
  if (event.Name === _main_macro_name()) {
    console.info('Remove Panel and Button for Macro Deactivation');
    macroReset();
  }
});

// Process Macro Removed
xapi.Event.Macros.Macro.Removed.on((event) => {
  if (event.Name === _main_macro_name()) {
    console.info('Remove Panel and Button for Macro Deletion');
    macroReset();
  }
});

// Init function
async function init() {
  console.info(`Report Issue Macro v${version}`);
  await removePanel(panelId, false);
  try {
    if (!processInputs()) throw new Error('Input Issues');
    const systemUnit = await xapi.Status.SystemUnit.get();
    sysInfo.version = systemUnit.Software.Version;
    // Determine device mode
    try {
      const mtrStatus = await xapi.Command.MicrosoftTeams.List();
      sysInfo.isRoomOS = !mtrStatus.Entry.some((i) => i.Status === 'Installed');
      if (!sysInfo.isRoomOS) { console.info('Device in Microsoft Mode'); }
    } catch (error) {
      // Device does not support MTR
      sysInfo.isRoomOS = true;
    }
    // Get System Name / Contact Name
    sysInfo.name = await xapi.Status.UserInterface.ContactInfo.Name.get();
    // Get System SN
    sysInfo.serial = systemUnit.Hardware.Module.SerialNumber;
    if (!sysInfo.name || sysInfo.name === '') {
      sysInfo.name = sysInfo.serial;
    }
    // Get codec platform
    sysInfo.platform = await xapi.status.get('SystemUnit.ProductPlatform');
    // if matches desk or board series, flag that paired nagivagtor needs to be checked
    if (sysInfo.platform.toLowerCase().includes('desk') || sysInfo.platform.toLowerCase().includes('board')) {
      if (o.logDetailed) console.debug('Device matches Desk or Board Series.');
      navigatorCheck = true;
    }
    // HTTP Client needed for sending outbound requests
    await xapi.Config.HttpClient.Mode.set('On');
    // Check for SNOW CMDB CI
    if (o.snowEnabled && o.snowCmdbLookup) {
      try {
        let result = await xapi.Command.HttpClient.Get({ Header: snowHeader, Url: `${snowCMDBUrl}?sysparm_limit=1&serial_number=${sysInfo.serial}` });
        result = JSON.parse(result.Body).result;
        // Validate CI Data
        if (result && result.length === 0) {
          if (o.logDetailed) console.warn('SNOW CI Not Found');
        }
        if (result && result.length === 1) {
          const [ciInfo] = result;
          sysInfo.cmdbCi = ciInfo.sys_id;
          if (o.logDetailed) console.info(`SNOW CI Found - ${ciInfo.name}`);
        }
      } catch (error) {
        console.error('unable to get SNOW CMDB CI');
        console.debug(error.message ? error.message : error);
      }
    }
    // Validate Button
    if (o.buttonEnabled) {
      await addPanel();
    }
    // Validate Debug Button
    if (o.debugButton) {
      await addDebugButton();
    }
  } catch (error) {
    console.error(error.message ? error.message : error);
    const Name = _main_macro_name();
    xapi.Command.Macros.Macro.Deactivate({ Name });
    console.error(`Macro ${Name} deactivated.`);
  }
}

init();

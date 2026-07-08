# Technical Reference

This document contains the implementation and configuration detail for the Report Issue macro.

## Quick Links

- [Prerequisites](#prerequisites)
- [Service Processing](#service-processing)
- [Deployment](#deployment)
- [Debugging](#debugging)
- [Configuration Reference](#configuration-reference)
- [Power BI Streaming Dataset](#power-bi-streaming-dataset)

## Prerequisites

The following items are needed depending on the services you enable.

### Webex Spaces

This integration supports either:

- Incoming Webhooks
- A Webex Bot

#### Incoming Webhook

- One or two Webex spaces with an [Incoming Webhook](https://apphub.webex.com/applications/incoming-webhooks-cisco-systems-38054-23307-75252) configured
- `webexWebhook` is the default webhook for Webex messages
- `webexReportWebhook` can be used to send Report Issue messages to a separate space

#### Webex Bot

- A Webex Bot created at [developer.webex.com](https://developer.webex.com/my-apps/new/bot)
- One or two Webex spaces with the bot added as a member
- `webexRoomId` is the default destination space
- `webexReportRoomId` can be used to send Report Issue messages to a separate space

Ways to get a Webex Room ID:

- Use the [List Rooms](https://developer.webex.com/docs/api/v1/rooms/list-rooms) API
- Add `astronaut@webex.bot` to the space and read the returned ID
- Start a 1:1 chat with `astronaut@webex.bot` and provide the space link or mention the space

### Microsoft Teams Channels

- One or two Microsoft Teams channels configured with an [Incoming Webhook with Workflows](https://support.microsoft.com/en-au/office/create-incoming-webhooks-with-workflows-for-microsoft-teams-8ae491c7-0394-4861-ba59-055e33f75498)
- `teamsWebhook` is the default webhook for Teams messages
- `teamsReportWebhook` can be used to send Report Issue messages to a separate channel

### ServiceNow

- A ServiceNow account with `sn_incident_write`
- Read access to CMDB and user records if lookup features are enabled
- Your ServiceNow instance hostname
- ServiceNow API key (recommended; see configuration steps below)
- Base64-encoded `username:password` credentials for basic authentication (legacy option)
- Optional default caller and default CI `sys_id` values

ServiceNow extra fields can be applied in this order:

- Globally in macro options
- At category level
- At issue level

#### ServiceNow Reporter Search

When `snowUserLookup` is enabled, the macro first checks ServiceNow for an exact reporter match using `snowUserField` and any configured `snowUserAppend` suffix. If that lookup finds no confident single match and `snowUserDisplayLookup` is enabled, the macro searches active `sys_user` records by display name, using `snowUserDisplayField` and `snowUserDisplaySearchMode`. Set `snowUserDisplaySearchMode` to `startsWith` to match display names that begin with the entered text, or `contains` to match display names that include the entered text anywhere in the name.

If display-name search returns one match, that user is selected automatically for the incident caller. If multiple matches are returned, the touch panel prompts the reporter to choose the correct user or refine the search. `snowUserDisplayLimit` controls how many selectable matches are shown before the refine option.

#### ServiceNow API Authentication

Refer to [this guide](https://www.servicenow.com/docs/r/platform-security/authentication/configure-api-key.html) for steps to configure header-based API key authentication.

- Include the `x-sn-apikey` header when creating the Inbound Authentication Profile
- Select `Table API` when creating the REST API Access Policy
- If you also need basic authentication, create an HTTP Basic Auth profile and assign it to the same REST API Access Policy

**Note:** Review any configuration changes with your ServiceNow administrator.

### HTTP JSON

- A remote service capable of receiving HTTP `POST` messages
- This can include tools such as Power BI or Loki

Example payload:

```json
{
  "timestamp": 1728875901099,
  "system": "West Ham United",
  "serial": "FOC1234567AB",
  "version": "ce11.22.0.18.55610ed00ae",
  "source": "call",
  "rating": 5,
  "rating_fmt": "Excellent",
  "destination": "spark:12345678901@cisco.webex.com",
  "type": "webex",
  "type_fmt": "Webex",
  "duration": 45,
  "duration_fmt": "45 seconds",
  "cause": "LocalDisconnect",
  "category": "",
  "category_fmt": "",
  "issue": "",
  "issue_fmt": "",
  "voluntary": 1
}
```

Notes:

- For Power BI, use `httpFormat: powerBi` so the timestamp is converted to `DateTime`
- For Loki, use `httpFormat: loki`

## Service Processing

The following table outlines how responses are processed for enabled services.

- `callEnabled` controls whether the post-call survey is shown
- `buttonEnabled` controls whether the Report Issue button is added to the touch panel

| Service | Option | Outcome |
| ---- | ---- | ---- |
| Webex | Any* | Message will be sent if comments are provided |
| Webex | Report Issue | Message will be sent, optionally to a separate space if defined by `webexReportRoomId` or `webexReportWebhook` |
| Webex | Survey - Excellent | Message will be sent if `webexLogExcellent` is enabled |
| Webex | Survey - Average/Poor | Message will be sent |
| MS Teams | Any* | Message will be sent if comments are provided |
| MS Teams | Report Issue | Message will be sent, optionally to a separate channel if defined by `teamsReportWebhook` |
| MS Teams | Survey - Excellent | Message will be sent if `teamsLogExcellent` is enabled |
| MS Teams | Survey - Average/Poor | Message will be sent |
| ServiceNow | Report Issue | `snowTicketReport` disabled: incident is raised. `snowTicketReport` enabled: incident is raised only if Raise Incident is selected on the panel |
| ServiceNow | Survey - Excellent | No incident is raised |
| ServiceNow | Survey - Average | `snowTicketCall` disabled: incident is raised if `snowRaiseAverage` is enabled. `snowTicketCall` enabled: incident is raised only if Raise Incident is selected on the panel |
| ServiceNow | Survey - Poor | `snowTicketCall` disabled: incident is raised. `snowTicketCall` enabled: incident is raised only if Raise Incident is selected on the panel |
| HTTP Server | Any | Output will always be sent to the HTTP server |

## Deployment

1. Download `ReportIssue.js` from this repository.
2. Update the macro options and enabled services at the top of the file.
3. Enable `debugButton` during initial setup so you can test without placing calls.
4. Upload the macro to your Cisco device and activate it.
5. Test the Report Issue button and, if enabled, the debug survey button.
6. If `callEnabled` is enabled, place a test call longer than `minDuration` and confirm the survey appears.
7. Validate that each enabled destination receives the expected output.

Macro console logs, especially debug-level logs, are useful when troubleshooting setup issues.

## Debugging

The macro includes an optional action button for testing the post-call survey flow without placing a real call.

Enable:

```javascript
debugButton: true,
```

## Configuration Reference

### Macro Options

#### Core

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| appName | string | `reportIssue` | App name used for logging and UI identifiers |
| widgetPrefix | string | `ri-` | Prefix used for UI widget and feedback identifiers |
| debugButton | bool | `false` | Enables the debug survey button for testing |

#### Call Parameters

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| callEnabled | bool | `false` | Enables post-call survey processing |
| minDuration | num | `10` | Minimum call duration in seconds before the survey is displayed |

#### Panel Parameters

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| panelEmoticons | bool | `true` | Shows emoticons on the panel |
| panelComments | bool | `true` | Shows the comments field on the panel |
| panelTips | bool | `true` | Shows helper text for category and issue selections |

#### User Selection

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| userParseEmail | bool | `false` | Parses and validates an email address in the user field |
| userSuggest | bool | `true` | Suggests entering a reporter when ticket creation is enabled |

#### Button Parameters

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| buttonEnabled | bool | `true` | Adds a Report Issue button to the UI |
| buttonLocation | string | `HomeScreenAndCallControls` | Button location. Options: `HomeScreen`, `HomeScreenAndCallControls`, `ControlPanel` |
| buttonPosition | num | `1` | Button order position |
| buttonIcon | string | `Concierge` | Icon used for Report Issue and Debug buttons |
| buttonColor | string | `#1170CF` | Button color |

#### Webex Messaging

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| webexEnabled | bool | `false` | Enables Webex space message logging |
| webexLogExcellent | bool | `false` | Logs excellent survey results to Webex |
| webexBotToken | string | `` | Webex bot token |
| webexRoomId | string | `` | Default Webex room ID |
| webexReportRoomId | string | `` | Optional room ID for Report Issue messages |
| webexWebhook | string | `` | Default Webex incoming webhook |
| webexReportWebhook | string | `` | Optional webhook for Report Issue messages |

#### Microsoft Teams Messaging

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| teamsEnabled | bool | `false` | Enables Microsoft Teams channel message logging |
| teamsLogExcellent | bool | `false` | Logs excellent survey results to Teams |
| teamsWebhook | string | `` | Default Teams webhook URL |
| teamsReportWebhook | string | `` | Optional webhook for Report Issue messages |

#### HTTP Server

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| httpEnabled | bool | `false` | Enables HTTP JSON POST delivery |
| httpUrl | string | `` | HTTP POST URL |
| httpAuth | bool | `false` | Adds a custom authentication header |
| httpHeader | string | `Authorization: XXXX` | Authentication header content |
| httpFormat | string | `none` | Optional formatting mode such as `none`, `loki`, or `powerBi` |

#### ServiceNow

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| snowEnabled | bool | `false` | Enables ServiceNow incident integration |
| snowTicketCall | bool | `true` | Shows the Raise Ticket option for post-call surveys |
| snowTicketReport | bool | `false` | Shows the Raise Ticket option for Report Issue submissions |
| snowRaiseAverage | bool | `false` | Raises incidents for average survey responses when `snowTicketCall` is disabled |
| snowUserLookup | bool | `true` | Performs ServiceNow user lookup when a reporter is entered |
| snowUserAppend | string | `` | Appends a suffix such as `@domain` before lookup |
| snowUserRequired | bool | `false` | Requires a user when raising a ticket |
| snowUserField | string | `user_name` | ServiceNow field used for user lookup |
| snowUserDisplayLookup | bool | `true` | Searches active ServiceNow users by display name when exact lookup is not confident |
| snowUserDisplayField | string | `name` | ServiceNow field used for fallback display-name search |
| snowUserDisplaySearchMode | string | `startsWith` | Display-name search mode. Options: `startsWith`, `contains` |
| snowUserDisplayLimit | num | `4` | Maximum selectable display-name matches shown before the refine option |
| snowInstance | string | `xxxx.service-now.com` | ServiceNow instance hostname |
| snowApiKey | string | `` | ServiceNow API key for header-based authentication |
| snowCredentials | string | `` | Base64-encoded basic auth credentials (legacy option) |
| snowCallerId | string | `` | Default caller `sys_id` |
| snowCmdbCi | string | `` | Default CMDB CI `sys_id` |
| snowCmdbLookup | bool | `true` | Looks up the device in ServiceNow by serial number |
| snowExtra | json | `{}` | Extra fields to pass to ServiceNow incidents |

#### Global Parameters

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| defaultSubmit | bool | `true` | Sends post-call results if the user does not explicitly submit before timeout |
| uploadLogsCallPoor | bool | `true` | Automatically uploads logs for poor ratings |
| uploadLogsCallAverage | bool | `false` | Automatically uploads logs for average ratings |
| uploadLogsReport | bool | `false` | Automatically uploads logs for Report Issue submissions |

#### Timeout Parameters

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| timeoutPanel | num | `20` | Timeout before the initial survey panel is dismissed, in seconds |
| timeoutPopup | num | `10` | Timeout before survey popups are dismissed, in seconds |

#### Logging Parameters

| Name | Type | Default | Description |
| ---- | ---- | ------- | ----------- |
| logDetailed | bool | `true` | Enables detailed logging |
| logUnknownResponses | bool | `false` | Logs unknown extension responses for troubleshooting |

### Language Options

| Name | Default | Description |
| ---- | ------- | ----------- |
| categoryTip | `Please select the most appropriate category` | Tip shown when selecting a category |
| issueTip | `Please select the most appropriate issue` | Tip shown when selecting an issue |
| buttonText | `Report Issue` | Button text shown on the touch panel |
| issuePrefix | `Report Issue` | Prefix used for report issue titles |
| feedbackPrefix | `Feedback` | Prefix used for post-call feedback titles |
| ticketTerm | `Incident` | Ticket terminology shown in the UI |
| userField | `Username` | Label used for the reporter field |
| userPlaceholder | `Please provide your username` | Placeholder shown for the reporter prompt |
| navigatorText | `Please complete Feedback Survey on the Touchpanel` | Alert shown on a Board when a Navigator is connected |

## Power BI Streaming Dataset

The following fields are required for a Power BI API streaming dataset.

| Value | Format | Comments |
| ---- | ---- | ---- |
| timestamp | `DateTime` | |
| system | `String` | |
| serial | `String` | |
| version | `String` | |
| source | `String` | |
| rating | `Number` | |
| rating_fmt | `String` | |
| destination | `String` | |
| type | `String` | |
| type_fmt | `String` | |
| duration | `Number` | |
| duration_fmt | `String` | |
| cause | `String` | |
| category | `String` | |
| category_fmt | `String` | |
| issue | `Number` | |
| issue_fmt | `String` | |
| comments | `String` | |
| reporter | `String` | |
| voluntary | `Number` | `0` for false, `1` for true |

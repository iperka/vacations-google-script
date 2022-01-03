/**
 * Vacations Google Script
 *
 * Allows to connect your Google calendar with the Vacations API.
 * It adds all your vacations to your account.
 *
 * IMPORTANT: If the example value ends with a slash be sure to add one to your value too.
 *
 * @version 1.0.1
 * @author Michael Beutler
 */

/**
 * If set to true the script will replace and update existing events.
 */
const REPLACE = true;

/**
 * Must be set to true in order to create events.
 */
const IS_ACTIVE = false;

/**
 * If set to false the calendar events will be called as the vacations API name suggests.
 * When assigned a string, all vacations will be named equally to the string.
 */
const OVERRIDE_SUMMARY = "Vacations";

/**
 * Your email to receive notifications.
 */
const REPORT_EMAIL = "MY_EMAIL";

/**
 * Google Calendar ID, this is usually your email address.
 */
const CALENDAR_ID = "MY_CALENDAR_ID";

/**
 * Auth0 Properties.
 */
const AUTH0_DOMAIN = "MY_AUTH0_DOMAIN";
const AUTH0_AUDIENCE = "MY_AUTH0_AUDIENCE";
const AUTH0_CLIENT_ID = "MY_CLIENT_ID";
const AUTH0_CLIENT_SECRET = "MY_CLIENT_SECRET";

const URL = "https://api.vacations.iperka.com/v1/vacations/";

function main() {
  let page = 0;
  let hasNextPage = false;
  do {
    const vacations = accessProtectedResource(URL + `?page=${page}`);
    if (vacations === null) {
      hasNextPage = false;
      continue;
    }
    Logger.log(
      `Fetching page ${page + 1} out of ${vacations.metadata.totalPages}...`
    );
    if (vacations.metadata.totalPages > page + 1) {
      hasNextPage = true;
    } else {
      hasNextPage = false;
    }

    const events = vacations.data.map((vacation) => ({
      id: toRFC2938(vacation.uuid),
      summary: OVERRIDE_SUMMARY ? OVERRIDE_SUMMARY : vacation.name,
      start: new Date(vacation.startDate),
      end: new Date(vacation.endDate),
      status: getStatus(vacation.status),
      vacation,
    }));

    events.forEach((e) => createEvent(e));
    page++;
  } while (hasNextPage);
}

function toRFC2938(id) {
  return id.replace(/[v-z\-]/g, "").toLowerCase();
}

function getStatus(status) {
  switch (status.toLowerCase()) {
    case "rejected":
    case "withdrawn":
      return "cancelled";
    case "requested":
      return "tentative";
    case "accepted":
    default:
      return "confirmed";
  }
}

/**
 * Attempts to create a new event with given event object.
 * When the event already exists and the global constant REPLACE is set to true
 * the event will get replaced.
 *
 * @param {Object} eventObject Event object.
 */
function createEvent(eventObject) {
  if (!eventObject) {
    return;
  }

  let allDayEvent = false;
  if (
    eventObject.start.getHours() === 20 &&
    eventObject.end.getHours() === 20
  ) {
    allDayEvent = true;
    eventObject.start.setDate(eventObject.start.getDate() + 1);
    eventObject.end.setDate(eventObject.end.getDate() + 2);
    eventObject.start.setHours(0, 0, 0, 0);
    eventObject.end.setHours(0, 0, 0, 0);
  } else {
    eventObject.start.setHours(eventObject.start.getHours() - 1);
    eventObject.end.setHours(eventObject.end.getHours() - 1);
  }

  var event = {
    id: eventObject.id,
    summary: eventObject.summary,
    location: eventObject.location,
    status: eventObject.status,
    start: allDayEvent
      ? { date: eventObject.start.toISOString().split("T")[0] }
      : {
          dateTime: eventObject.start.toISOString(),
        },
    end: allDayEvent
      ? { date: eventObject.end.toISOString().split("T")[0] }
      : {
          dateTime: eventObject.end.toISOString(),
        },
    description: eventObject.description,
    source: {
      url: URL + eventObject.vacation.uuid,
      title: "Vacations API",
    },
    creator: {
      id: eventObject.vacation.owner,
    },
    reminders: {
      useDefault: false,
      overrides: [],
    },
  };

  try {
    // Check if IS_ACTIVE
    if (IS_ACTIVE) {
      Calendar.Events.insert(event, CALENDAR_ID);
    }
    Logger.log(
      `Created event '${
        eventObject.summary
      }' at ${eventObject.start.toISOString()} - ${eventObject.end.toISOString()}.`
    );
  } catch (e) {
    if (e.message.includes("The requested identifier already exists")) {
      Logger.log(`Event with id '${eventObject.id}' already exists.`);
      if (REPLACE && IS_ACTIVE) {
        Calendar.Events.update(event, CALENDAR_ID, eventObject.id);
        Logger.log(
          `Updated event '${
            eventObject.summary
          }' at ${eventObject.start.toISOString()} - ${eventObject.end.toISOString()}.`
        );
      }
    } else {
      Logger.log(`An unexpected error occurred: ${e.message}`);
    }
  }
}

/**
 * Attempts to access a non-Google API using a constructed service
 * object.
 *
 * If your add-on needs access to non-Google APIs that require OAuth,
 * you need to implement this method. You can use the OAuth1 and
 * OAuth2 Apps Script libraries to help implement it.
 *
 * @param {String} url         The URL to access.
 * @param {String} method_opt  The HTTP method. Defaults to GET.
 * @param {Object} headers_opt The HTTP headers. Defaults to an empty
 *                             object. The Authorization field is added
 *                             to the headers in this method.
 * @return {HttpResponse} the result from the UrlFetchApp.fetch() call.
 */
function accessProtectedResource(url, method_opt, headers_opt) {
  var service = getOAuthService();
  var maybeAuthorized = service.hasAccess();

  if (maybeAuthorized) {
    // A token is present, but it may be expired or invalid. Make a
    // request and check the response code to be sure.

    // Make the UrlFetch request and return the result.
    var accessToken = service.getAccessToken();
    var method = method_opt || "get";
    var headers = headers_opt || {};

    headers["Authorization"] = Utilities.formatString("Bearer %s", accessToken);
    var resp = UrlFetchApp.fetch(url, {
      headers: headers,
      method: method,
      muteHttpExceptions: true, // Prevents thrown HTTP exceptions.
    });

    var code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      return JSON.parse(resp.getContentText("utf-8")); // Success
    } else if (code == 401 || code == 403) {
      // Not fully authorized for this action.
      maybeAuthorized = false;
    } else {
      // Handle other response codes by logging them and throwing an
      // exception.
      console.error(
        "Backend server error (%s): %s",
        code.toString(),
        resp.getContentText("utf-8")
      );
      throw "Backend server error: " + code;
    }
  }

  if (!maybeAuthorized) {
    // Invoke the authorization flow using the default authorization
    // prompt card.
    Logger.log(
      "Open the following URL and re-run the script: %s",
      service.getAuthorizationUrl()
    );
    GmailApp.sendEmail(
      REPORT_EMAIL,
      `Authentication required.`,
      `Vacations syncronization failed due to missing authentication.\nPlease visit the following link to authenticate.\n${service.getAuthorizationUrl()}`,
      {
        noReply: true,
        htmlBody: `Vacations syncronization failed due to missing authentication.<br />Please visit the following link to authenticate.<br /><a href="${service.getAuthorizationUrl()}">Authenticate</a>`,
      }
    );
    return null;
  }
}

/**
 * Create a new OAuth service to facilitate accessing an API.
 * This example assumes there is a single service that the add-on needs to
 * access. Its name is used when persisting the authorized token, so ensure
 * it is unique within the scope of the property store. You must set the
 * client secret and client ID, which are obtained when registering your
 * add-on with the API.
 *
 * See the Apps Script OAuth2 Library documentation for more
 * information:
 *   https://github.com/googlesamples/apps-script-oauth2#1-create-the-oauth2-service
 *
 *  @return A configured OAuth2 service object.
 */
function getOAuthService() {
  return (
    OAuth2.createService("Auth0")
      .setAuthorizationBaseUrl(AUTH0_DOMAIN + "/authorize")
      .setTokenUrl(AUTH0_DOMAIN + "/oauth/token")
      .setClientId(AUTH0_CLIENT_ID)
      .setClientSecret(AUTH0_CLIENT_SECRET)
      .setScope("vacations:read")
      .setCallbackFunction("authCallback")
      .setCache(CacheService.getUserCache())
      .setPropertyStore(PropertiesService.getUserProperties())

      // Below are Auth0-specific OAuth2 parameters.
      .setParam("audience", AUTH0_AUDIENCE)
      .setParam("response_type", "code")
      .setParam("response_mode", "query")
  );
}

/**
 * Boilerplate code to determine if a request is authorized and returns
 * a corresponding HTML message. When the user completes the OAuth2 flow
 * on the service provider's website, this function is invoked from the
 * service. In order for authorization to succeed you must make sure that
 * the service knows how to call this function by setting the correct
 * redirect URL.
 *
 * The redirect URL to enter is:
 * https://script.google.com/macros/d/<Apps Script ID>/usercallback
 *
 * See the Apps Script OAuth2 Library documentation for more
 * information:
 *   https://github.com/googlesamples/apps-script-oauth2#1-create-the-oauth2-service
 *
 *  @param {Object} callbackRequest The request data received from the
 *                  callback function. Pass it to the service's
 *                  handleCallback() method to complete the
 *                  authorization process.
 *  @return {HtmlOutput} a success or denied HTML message to display to
 *          the user. Also sets a timer to close the window
 *          automatically.
 */
function authCallback(callbackRequest) {
  console.log(callbackRequest);
  var authorized = getOAuthService().handleCallback(callbackRequest);
  if (authorized) {
    return HtmlService.createHtmlOutput(
      "Success! <script>setTimeout(function() { top.window.close() }, 1);</script>"
    );
  } else {
    return HtmlService.createHtmlOutput("Denied");
  }
}

/**
 * Unauthorizes the non-Google service. This is useful for OAuth
 * development/testing.  Run this method (Run > resetOAuth in the script
 * editor) to reset OAuth to re-prompt the user for OAuth.
 */
function resetOAuth() {
  getOAuthService().reset();
}

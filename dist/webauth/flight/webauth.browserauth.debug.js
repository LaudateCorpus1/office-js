var BrowserAuth =
/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ({

/***/ "./packages/Microsoft.Office.WebAuth.BrowserAuth/lib/api.js":
/*!******************************************************************!*\
  !*** ./packages/Microsoft.Office.WebAuth.BrowserAuth/lib/api.js ***!
  \******************************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.logUserAction = logUserAction;
exports.logActivity = logActivity;
exports.sendTelemetryEvent = sendTelemetryEvent;
exports.sendActivityEvent = sendActivityEvent;
exports.sendOtelEvent = sendOtelEvent;
exports.sendUserActionEvent = sendUserActionEvent;
exports.addNamespaceMapping = addNamespaceMapping;
exports.setEnabledState = setEnabledState;
exports.shutdown = shutdown;
exports.registerEventHandler = registerEventHandler;

var telemetryEnabled = false;
var events = [];
var eventHandler;
var numberOfDroppedEvents = 0;
var maxQueueSize = 20000;
var unknownStr = 'Unknown'; // Primary consumer public API

// ========================================================================================================================
// Call LogUserAction for logging a user action to Otel.
// This is similar to the bSqm actions that used to be logged earlier (deprecated now).
// Make sure you read the documentation below for userActionName and the Kusto table name implications.
// userActionName: Name of the user action, this should come from your app's commands,
//     for example: OneNoteCommands in office-online-ui\packages\onenote-online-ux\src\store\OneNoteCommands.ts (https://office.visualstudio.com/OC/_git/office-online-ui?path=%2Fpackages%2Fonenote-online-ux%2Fsrc%2Fstore%2FOneNoteCommands.ts&version=GBmaster)
//     Note that the userActionName will be the name of your table in Aria Kusto. So if 'ABC' is passed in for userActionName, the table in Kusto will be called Office_OneNote_Online_UserAction_ABC (or generically speaking Office_{AppName}_Online_UserAction_ABC )
//     Look at Kusto connection https://kusto.aria.microsoft.com and databases Office Word Online or Office OneNote Online, etc. and look at *UserAction* tables.
// success: Status of the user action (success is true, failure is false).
// parentNameStr: parent surface of the user action (example, tabView, tabHelp, Layout, etc).
// inputMethod: how the user action was performed (for example, via keyboard, or mouse, touch, etc.)
//             See the enum in /packages/app-commanding-ui/src/UISurfaces/controls/InputMethod.ts
//             Pass in this param as:  InputMethod.Keyboard.toString() instead of passing in "Keyboard"
// uiLocation: the surface where the user action was initiated from (example, ribbon, FileMenu, TellMe, etc).
//             See enum in /packages/app-commanding-ui/src/UISurfaces/controls/UILocation.ts
//             Pass in this param as:  UILocation.SingleLineRibbon.toString() instead of passing in "SingleLineRibbon"
// durationMsec: the time taken by the action (if relevant to the action)
// dataFieldArr: These are custom fields that you may want to add for your user action.
//               Example: InsertTable action may log custom data fields such as rowSize and colSize of the table inserted.
//                      Or in Excel, a cell related action may log the x and y coordinates of the cell.
// Note that things such as sessionID, data center, etc will be added to all user action logs.
function logUserAction(userActionName, success, parentNameStr, inputMethod, uiLocation, durationMsec, dataFieldArr) {
  if (success === void 0) {
    success = true;
  }

  if (parentNameStr === void 0) {
    parentNameStr = unknownStr;
  }

  if (inputMethod === void 0) {
    inputMethod = unknownStr;
  }

  if (uiLocation === void 0) {
    uiLocation = unknownStr;
  }

  if (durationMsec === void 0) {
    durationMsec = 0;
  }

  if (dataFieldArr === void 0) {
    dataFieldArr = [];
  }

  // passing null for 'name' field, which is the event table name. We will determine that in sendUserAction in full\api.ts as there we know what app we are, and hence what the event table name is
  sendUserActionEvent({
    name: null,
    actionName: userActionName,
    commandSurface: uiLocation,
    parentName: parentNameStr,
    triggerMethod: inputMethod,
    durationMs: durationMsec,
    succeeded: success,
    dataFields: dataFieldArr
  });
}

// Call logActivity for logging an activity to Otel.
// This will be logged under Office {App} Online Data tenant
// For example, if your activity name is "ABC",
// it will go to a table called "Office_Word_Online_Data_Activity_ABC" for Word or "Office_OneNote_Online_Data_Activity_ABC" for OneNote.
// activityName: name of activity being logged
// success: Status of the activity (success is true, failure is false).
// durationMsec: the time taken by the action (if relevant to the action)
// dataFieldArr: These are custom fields that you may want to add for your activity, and will be added as columns to the activity table.
//               Example: dataFields has typingSpeedPerSec (integer) and dayOfWeek (string) in it, the activity table for this particular activity will contain these two custom fields.
// Note that things such as sessionID, data center, etc will be added to all user action logs.
function logActivity(activityName, success, durationMsec, dataFieldArr) {
  if (success === void 0) {
    success = true;
  }

  if (durationMsec === void 0) {
    durationMsec = 0;
  }

  if (dataFieldArr === void 0) {
    dataFieldArr = [];
  }

  sendActivityEvent({
    name: activityName,
    succeeded: success,
    durationMs: durationMsec,
    dataFields: dataFieldArr
  });
}

function sendTelemetryEvent(event) {
  raiseEvent({
    kind: 'event',
    event: event,
    timestamp: new Date().getTime()
  });
}

function sendActivityEvent(event) {
  raiseEvent({
    kind: 'activity',
    event: event,
    timestamp: new Date().getTime()
  });
}

function sendOtelEvent(event) {
  raiseEvent({
    kind: 'otel',
    event: event
  });
}

function sendUserActionEvent(event) {
  raiseEvent({
    kind: 'action',
    event: event,
    timestamp: new Date().getTime()
  });
}

function addNamespaceMapping(namespace, ariaTenantToken) {
  raiseEvent({
    kind: 'addNamespaceMapping',
    namespace: namespace,
    ariaTenantToken: ariaTenantToken
  });
}

function setEnabledState(enabled) {
  telemetryEnabled = enabled; // If the caller disables the queue, be sure to drop all of the outstanding events.
  // This can happen in cases where the slice with event processor functionality failed to load.

  if (!telemetryEnabled) {
    events = [];
  }
}

function shutdown() {
  raiseEvent({
    kind: 'shutdown'
  });
  return events.length + numberOfDroppedEvents;
}

function registerEventHandler(handler) {
  eventHandler = handler; // Then go through the queue and process the events in the order in which they were received
  // VSO.2533164: Push batch event processing to otelFull and add a lightweight queue

  events.forEach(function (event) {
    return raiseEvent(event);
  });
  events = [];
}

function raiseEvent(event) {
  if (!telemetryEnabled) {
    return;
  }

  if (eventHandler) {
    eventHandler(event);
  } else {
    if (events.length <= maxQueueSize) {
      events.push(event);
    } else {
      numberOfDroppedEvents += 1;
    }
  }
}


/***/ }),

/***/ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/BrowserAuth.ts":
/*!******************************************************************************!*\
  !*** ./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/BrowserAuth.ts ***!
  \******************************************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/**
 * Copyright (c) Microsoft. All rights reserved.
 */
/// <reference path="./Definitions/IBrowserAuthConfig.d.ts" />
/// <reference path="./Definitions/IBrowserAuthResult.d.ts" />
// Above references are needed for ts-node
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var Constants_1 = __webpack_require__(/*! ./Utils/Constants */ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/Constants.ts");
exports.Constants = Constants_1.Constants;
var ExtractUtils_1 = __webpack_require__(/*! ./Utils/ExtractUtils */ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/ExtractUtils.ts");
exports.ExtractUtils = ExtractUtils_1.ExtractUtils;
var LoggingUtils_1 = __webpack_require__(/*! ./Utils/LoggingUtils */ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/LoggingUtils.ts");
exports.LoggingUtils = LoggingUtils_1.LoggingUtils;
var ValidateUtils_1 = __webpack_require__(/*! ./Utils/ValidateUtils */ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/ValidateUtils.ts");
exports.ValidateUtils = ValidateUtils_1.ValidateUtils;
var TimerUtils_1 = __webpack_require__(/*! ./Utils/TimerUtils */ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/TimerUtils.ts");
// This is the aria token used for Otel telemetry event ingestion.
var api_js_1 = __webpack_require__(/*! ../lib/api.js */ "./packages/Microsoft.Office.WebAuth.BrowserAuth/lib/api.js");
api_js_1.addNamespaceMapping('Office.Identity.WebAuth.BrowserAuth', '5c65bbc4edbf480d9637ace04d62bd98-12844893-8ab9-4dde-b850-5612cb12e0f2-7822');
var msalInstance;
exports.msalInstance = msalInstance;
var applications;
exports.applications = applications;
var authConfig;
/**
 * Load auth code module
 * @param configurations - auth configs
 * @param correlationId - the correlation ID exists under the caller's context, the same correlation ID will passed on to MSAL.js for unified experience
 * @returns the {@link IBrowserAuthLoadResult} object
 */
function Load(configurations, correlationId) {
    var timerClock = TimerUtils_1.TimerUtils.timer();
    authConfig = configurations;
    if (authConfig.enableConsoleLogging) {
        // Turn on if explicitly authConfig.enableConsoleLogging is true
        LoggingUtils_1.LoggingUtils.enableConsoleLogging = authConfig.enableConsoleLogging;
    }
    if (authConfig.enableUpnCheck) {
        // Turn on if explicitly authConfig.enableUpnCheck is true
        ValidateUtils_1.ValidateUtils.enableUpnCheck = authConfig.enableUpnCheck;
    }
    // Implicitly swap login.windows.net to login.microsoftonline.com
    if (authConfig.authority) {
        authConfig.authority = authConfig.authority.replace(Constants_1.Constants.Authority.ProdLegacy, Constants_1.Constants.Authority.Prod);
    }
    var isMsa = authConfig.idp.toLowerCase() === Constants_1.Constants.IdentityProvider.Msa.toLowerCase();
    if (isMsa) {
        // There is no guarantee that our callers will send MSAL.js supported authority to us,
        // hard code the authority to be login.microsoftonline.com since there is no PPE tenant for MSA.
        authConfig.authority = Constants_1.Constants.Authority.Prod + Constants_1.Constants.Authority.MsaSuffix;
    }
    else {
        if (!authConfig.authority) {
            authConfig.authority = Constants_1.Constants.Authority.Prod + Constants_1.Constants.Authority.AadSuffix;
        }
        // Add common suffix if it doesn't have
        if (authConfig.authority.indexOf(Constants_1.Constants.Authority.AadSuffix) < 0) {
            if (authConfig.authority.charAt(authConfig.authority.length - 1) == "/")
                authConfig.authority += Constants_1.Constants.Authority.AadSuffix;
            else
                authConfig.authority += "/" + Constants_1.Constants.Authority.AadSuffix;
        }
    }
    return MsalWrapper((authConfig.timeoutToLoad) ? authConfig.timeoutToLoad : 2000).then(function (msalInstance) {
        exports.applications = applications = new Array();
        var defaultAppId;
        for (var _i = 0, _a = authConfig.appIds; _i < _a.length; _i++) {
            var appId = _a[_i];
            LoggingUtils_1.LoggingUtils.log("[Load->MsalWrapper->MsalInit] " + appId + ": " + authConfig.idp);
            if (!appId || !ValidateUtils_1.ValidateUtils.isGuid(appId)) {
                continue;
            }
            var application = new msalInstance.PublicClientApplication({
                auth: {
                    clientId: appId,
                    authority: authConfig.authority,
                    redirectUri: (authConfig.redirectUri) ? authConfig.redirectUri.split("?")[0] : location.href.split("?")[0],
                    navigateToLoginRequestUrl: (authConfig.navigateToLoginRequestUrl) ? authConfig.navigateToLoginRequestUrl : true,
                },
                cache: {
                    cacheLocation: 'localStorage',
                    // Store auth state in cookies can make the request too big and fail the request sometimes, need to keep it as false.
                    storeAuthStateInCookie: false
                },
                system: {
                    loadFrameTimeout: (authConfig.timeout) ? authConfig.timeout : 6000,
                },
            });
            var entry = { applicationId: appId, application: application };
            applications.push(entry);
            // Keep first valid app for pre-fetch
            if (!defaultAppId)
                defaultAppId = appId;
        }
        ;
        if (!defaultAppId)
            return Promise.reject(LogTelemetryDataFieldsForLoad(false, correlationId, applications, timerClock));
        LoggingUtils_1.LoggingUtils.log("[Load->SsoSilentWrapper] " + defaultAppId + ": " + authConfig.idp);
        return SsoSilentWrapper(defaultAppId, authConfig.upn, isMsa ? Constants_1.Constants.Authority.MsaSuffix : Constants_1.Constants.Authority.AadSuffix, correlationId).then(function (ssoResult) {
            LoggingUtils_1.LoggingUtils.log("[Load->SsoSilentWrapper] " + defaultAppId + ": success: " + ssoResult.idTokenClaims["aud"]);
            return LogTelemetryDataFieldsForLoad(true, correlationId, applications, timerClock);
        }).catch(function (ssoError) {
            LoggingUtils_1.LoggingUtils.log("[Load->SsoSilentWrapper] " + defaultAppId + ": error: [" + ssoError.errorCode + "]" + ssoError.errorMessage);
            return LogTelemetryDataFieldsForLoad(true, correlationId, applications, timerClock, ssoError.errorCode, ssoError.errorMessage);
        });
    }).catch(function () {
        LoggingUtils_1.LoggingUtils.log("[Load->MsalWrapper] msal is not loaded");
        return Promise.reject(LogTelemetryDataFieldsForLoad(false, correlationId, applications, timerClock));
    });
}
exports.Load = Load;
/**
 * Acquire an access token by given target
 * @param target - resource for V1 token, scope for V2 token
 * @param applicationId - the application ID which needs access token
 * @param correlationId - the same correlation ID exists under the caller's context, the same correlation ID will passed on to MSAL.js for unified experience
 * @param login - If true, shows a login dialog. If false, skips login. TODO: Remove after migration in WAC
 * @param popup - If true, popsup a dialog for interactive flow. If false, acquires token silently.
 * @param forThirdParty - If true, treats the caller as third-party and avoids sending PII telemetry.
 * @param claims - Claims from AAD to be used in scenarios such as credentials change or MFA
 * @returns {Promise.<IBrowserAuthResult>} - a promise that is fulfilled when this function has completed, or rejected if an error was raised. Returns the {@link IBrowserAuthResult} object
 */
function GetToken(target, applicationId, correlationId, login, popup, forThirdParty, claims) {
    var _this = this;
    if (forThirdParty === void 0) { forThirdParty = false; }
    var timerClock = TimerUtils_1.TimerUtils.timer();
    var isPopup = login || popup; //TODO:Need to deprecate login because one popup can control. Deprecate after WAC migration.
    var is3rdPartyCookie = false;
    var application = GetApplication(applicationId);
    var result = {};
    var isMsa = authConfig.idp.toLowerCase() === Constants_1.Constants.IdentityProvider.Msa.toLowerCase();
    // Wrong format of correlation ID or blank are not valid in MSAL.js
    // With an invalid correlation ID in the request, the access token acquiring request will be rejected by MSAL.js with exceptions.
    // Correlation ID will be set to undefined in those cases and MSAL.js will generate a new correlation ID if it is undefined.
    if (!correlationId || !ValidateUtils_1.ValidateUtils.isGuid(correlationId)) {
        correlationId = undefined;
    }
    if (!target) {
        result.ErrorCode = 'missing_target';
        result.ErrorMessage = 'The provided target for BrowserAuth.GetToken is null, blank or empty';
        return Promise.reject(LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, undefined));
    }
    var scopes = [ValidateUtils_1.ValidateUtils.getScope(target)];
    if (!applicationId || !ValidateUtils_1.ValidateUtils.isGuid(applicationId)) {
        result.ErrorCode = 'invalid_application_ID';
        result.ErrorMessage = 'The provided application ID for BrowserAuth.GetToken is null, blank, empty or with invalid format';
        return Promise.reject(LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, scopes));
    }
    if (application === undefined || !application) {
        result.ErrorCode = 'missing_msal';
        result.ErrorMessage = 'BrowserAuth can\'t find msal instance';
        return Promise.reject(LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, undefined));
    }
    var acquireToken = function (tokenParams) {
        return new Promise(function (resolve, reject) { return __awaiter(_this, void 0, void 0, function () {
            var tokenRequest;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken] " + applicationId + ": popup:[" + isPopup + "]");
                        // Reset result
                        result = {};
                        tokenRequest = (isPopup)
                            ? application.acquireTokenPopup(tokenParams)
                            : application.acquireTokenSilent(tokenParams);
                        return [4 /*yield*/, tokenRequest.then(function (authResponse) {
                                result.MsalResult = authResponse;
                                if (authResponse.accessToken) {
                                    if (ValidateUtils_1.ValidateUtils.upnMatchesUpnFromIdToken(authResponse.idToken, authConfig)
                                        && (forThirdParty || isMsa || ValidateUtils_1.ValidateUtils.upnMatchesUpnFromAccessToken(authResponse.accessToken, authConfig))) {
                                        result.Token = authResponse.accessToken;
                                        LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken] " + applicationId + ": success: " + result.Token);
                                        return resolve(result);
                                    }
                                    else {
                                        result.ErrorCode = "upn_mismatch";
                                        result.ErrorMessage = "upn doesn't match with given upn in config";
                                        LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken] " + applicationId + ": error: [" + result.ErrorCode + "]" + result.ErrorMessage);
                                        return reject(result);
                                    }
                                }
                                result.ErrorCode = "no_tokens_found";
                                result.ErrorMessage = "Access token doesn't exist in auth response";
                                LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken] " + applicationId + ": error: [" + result.ErrorCode + "]" + result.ErrorMessage);
                                return reject(result);
                            }).catch(function (error) {
                                result.MsalResult = error;
                                result.ErrorCode = error.errorCode;
                                result.ErrorMessage = error.errorMessage;
                                LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken] " + applicationId + ": error: [" + result.ErrorCode + "]" + result.ErrorMessage);
                                if (!isPopup && is3rdPartyCookie) {
                                    isPopup = true;
                                    LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken] " + applicationId + ": retry: popup:[" + isPopup + "]");
                                    acquireToken(tokenParams).then(function (result) { return resolve(result); }).catch(function (result) { return reject(result); });
                                }
                                else {
                                    return reject(result);
                                }
                            })];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        }); });
    };
    var tokenParams = {
        scopes: scopes,
        account: { username: authConfig.upn },
        loginHint: authConfig.upn,
        extraQueryParameters: { domain_hint: isMsa ? Constants_1.Constants.Authority.MsaSuffix : Constants_1.Constants.Authority.AadSuffix },
        correlationId: correlationId,
        claims: claims
    };
    return acquireToken(tokenParams).catch(function (error) { return __awaiter(_this, void 0, void 0, function () {
        var isSuccess;
        var _this = this;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    isSuccess = false;
                    if (!(error.ErrorCode === "no_tokens_found")) return [3 /*break*/, 2];
                    LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken] " + applicationId + ": no_tokens_found");
                    return [4 /*yield*/, SsoSilentWrapper(applicationId, authConfig.upn, isMsa ? Constants_1.Constants.Authority.MsaSuffix : Constants_1.Constants.Authority.AadSuffix, correlationId).then(function () { return __awaiter(_this, void 0, void 0, function () {
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken->SsoSilentWrapper] " + applicationId + ": success");
                                        return [4 /*yield*/, acquireToken(tokenParams).then(function (tokenResult) {
                                                isSuccess = true;
                                                result = tokenResult;
                                            }).catch(function (tokenResult) {
                                                result = tokenResult;
                                            })];
                                    case 1:
                                        _a.sent();
                                        return [2 /*return*/];
                                }
                            });
                        }); }).catch(function (ssoError) { return __awaiter(_this, void 0, void 0, function () {
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        result.ErrorCode = ssoError.errorCode;
                                        result.ErrorMessage = ssoError.errorMessage;
                                        LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken->SsoSilentWrapper] " + applicationId + ": error: [" + result.ErrorCode + "]" + result.ErrorMessage);
                                        if (!(ssoError.errorCode === 'login_required' && (ssoError.errorMessage.startsWith('AADSTS50058') || ssoError.errorMessage.startsWith('Silent authentication was denied')) && authConfig.autoPopup)) return [3 /*break*/, 2];
                                        is3rdPartyCookie = true;
                                        return [4 /*yield*/, acquireToken(tokenParams).then(function (tokenResult) {
                                                isSuccess = true;
                                                result = tokenResult;
                                            }).catch(function (tokenResult) {
                                                LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken->SsoSilentWrapper->Popup] " + applicationId + ": error: [" + result.ErrorCode + "]" + result.ErrorMessage);
                                                result = tokenResult;
                                            })];
                                    case 1:
                                        _a.sent();
                                        _a.label = 2;
                                    case 2: return [2 /*return*/];
                                }
                            });
                        }); })];
                case 1:
                    _a.sent();
                    _a.label = 2;
                case 2:
                    LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken->isSuccess] " + applicationId + ": [" + isSuccess + "]");
                    result = LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, scopes);
                    return [2 /*return*/, isSuccess ? Promise.resolve(result) : Promise.reject(result)];
            }
        });
    }); }).then(function (result) {
        LoggingUtils_1.LoggingUtils.log("[GetToken->acquireToken] " + applicationId + ": success");
        return Promise.resolve(LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, scopes));
    });
}
exports.GetToken = GetToken;
/**
 * Acquire an access token by given target with load
 * @param target - resource for V1 token, scope for V2 token
 * @param applicationId - the application ID which needs access token
 * @param isMSA - Identity Type is MSA or not
 * @param isPPE - PPE or Prod Authority
 * @returns {Promise.<IBrowserAuthResult>} - a promise that is fulfilled when this function has completed, or rejected if an error was raised. Returns the {@link IBrowserAuthResult} object
 */
function GetTokenOnce(target, applicationId, isMSA, isPPE) {
    if (isMSA === void 0) { isMSA = false; }
    if (isPPE === void 0) { isPPE = false; }
    return Load({
        idp: isMSA ? Constants_1.Constants.IdentityProvider.Msa.toLowerCase() : Constants_1.Constants.IdentityProvider.Aad.toLowerCase(),
        appIds: [applicationId],
        authority: isPPE ? Constants_1.Constants.Authority.Ppe : Constants_1.Constants.Authority.Prod,
        upn: "",
        autoPopup: true,
        enableConsoleLogging: true
    }).then(function () {
        // GetToken
        return GetToken(target, applicationId, undefined, true);
    }).catch(function (loadFailure) {
        return Promise.reject(loadFailure);
    });
}
exports.GetTokenOnce = GetTokenOnce;
/**
 * Wrapper of msal object
 */
function MsalWrapper(max) {
    if (max === void 0) { max = 2000; }
    return new Promise(function (resolve, reject) {
        var msalExisted = function () {
            try {
                return (msal !== undefined && !!msal);
            }
            catch (error) {
                return false;
            }
        };
        var waited = 0;
        var interval = 500;
        var delay = setInterval(function () {
            waited += interval;
            if (msalExisted()) {
                clearInterval(delay);
                return resolve(msal);
            }
            else if (waited > max) {
                clearInterval(delay);
                return reject(null);
            }
        }, interval);
    });
}
exports.MsalWrapper = MsalWrapper;
/**
 * Wrapper of ssoSilent for easy unit testing
 */
function SsoSilentWrapper(appId, upn, domain_hint, correlationId) {
    var application = GetApplication(appId);
    return application.ssoSilent({
        loginHint: upn,
        extraQueryParameters: { domain_hint: domain_hint },
        correlationId: correlationId,
    });
}
exports.SsoSilentWrapper = SsoSilentWrapper;
/**
 * Construct the Msal.PublicClientApplication instance for V2 endpoint calls.
 * @param applicationId - the application ID used to find or construct the MSAL instance.
 * @returns the Msal.PublicClientApplication instance to make calls to V2 endpoint.
 */
function GetApplication(applicationId) {
    var application;
    if (applications) {
        applications.some(function (value) {
            if (applicationId && value.applicationId && applicationId.toUpperCase() === value.applicationId.toUpperCase()) {
                application = value.application;
                LoggingUtils_1.LoggingUtils.log("[GetApplication] " + applicationId + ": Found");
                return true;
            }
            LoggingUtils_1.LoggingUtils.log("[GetApplication] " + applicationId + ": Not Found");
            return false;
        });
    }
    if (!application && !!msalInstance) {
        LoggingUtils_1.LoggingUtils.log("[GetApplication] " + applicationId + ": Add");
        application = new msalInstance.PublicClientApplication({
            auth: {
                clientId: applicationId,
                authority: authConfig.authority,
                redirectUri: (authConfig.redirectUri) ? authConfig.redirectUri.split("?")[0] : location.href.split("?")[0],
                navigateToLoginRequestUrl: (authConfig.navigateToLoginRequestUrl) ? authConfig.navigateToLoginRequestUrl : true,
            },
            cache: {
                cacheLocation: 'localStorage',
                // Store auth state in cookies can make the request too big and fail the request sometimes, need to keep it as false.
                storeAuthStateInCookie: false
            },
            system: {
                loadFrameTimeout: (authConfig.timeout) ? authConfig.timeout : 6000,
            },
        });
        var entry = { applicationId: applicationId, application: application };
        applications.push(entry);
    }
    return application;
}
exports.GetApplication = GetApplication;
/**
 * Acquire an Id token by given upn/target, and check if the upn matches context
 * @param correlationId - the same correlation ID exists under the caller's context, the same correlation ID will passed on to MSAL.js for unified experience
 * @returns {Promise.<IBrowserAuthUpnResult>} - a promise that is fulfilled when this function has completed, or rejected if an error was raised. Returns the {@link IBrowserAuthUpnResult} object
 */
function CheckUpnMatchIdToken(applicationId, correlationId) {
    var timerClock = TimerUtils_1.TimerUtils.timer();
    // Wrong format of correlation ID or blank are not valid in MSAL.js
    // With an invalid correlation ID in the request, the access token acquiring request will be rejected by MSAL.js with exceptions.
    // Correlation ID will be set to undefined in those cases and MSAL.js will generate a new correlation ID if it is undefined.
    if (!correlationId || !ValidateUtils_1.ValidateUtils.isGuid(correlationId)) {
        correlationId = undefined;
    }
    var result = {};
    if (!applicationId || !ValidateUtils_1.ValidateUtils.isGuid(applicationId)) {
        result.ErrorCode = 'invalid_application_ID';
        result.ErrorMessage = 'The provided application ID for BrowserAuth.CheckUpnMatchIdToken is null, blank, empty or with invalid format';
        LogTelemetryDataFieldsForCheckUpn(result, correlationId, applicationId, timerClock, undefined, result.ErrorCode, result.ErrorMessage);
        return Promise.reject(result);
    }
    var application = GetApplication(applicationId);
    if (application === undefined || !application) {
        result.ErrorCode = 'missing_msal';
        result.ErrorMessage = 'BrowserAuth can\'t find msal instance';
        return Promise.reject();
    }
    var scopes = [applicationId]; /*It will acquire id token instead of access token if resource/scope is clientId*/
    return new Promise(function (resolve, reject) {
        var currentAccount = application.getAccountByUsername(authConfig.upn);
        LoggingUtils_1.LoggingUtils.log("[CheckUpnMatchIdToken] " + applicationId + ": Calls acquireTokenSilent");
        application.acquireTokenSilent({
            scopes: scopes,
            account: currentAccount,
            correlationId: correlationId,
        }).then(function (authResponse) {
            result.IsUpnMatch = ValidateUtils_1.ValidateUtils.upnMatchesUpnFromIdToken(authResponse.idToken, authConfig);
            if (result.IsUpnMatch) {
                LogTelemetryDataFieldsForCheckUpn(result, correlationId, applicationId, timerClock, scopes);
                return resolve(result);
            }
        }).catch(function (error) {
            result.ErrorCode = error.errorCode;
            result.ErrorMessage = error.errorMessage;
            LoggingUtils_1.LoggingUtils.log("[CheckUpnMatchIdToken] " + applicationId + ": error: " + result.ErrorMessage);
            result.IsUpnMatch = false;
            LogTelemetryDataFieldsForCheckUpn(result, correlationId, applicationId, timerClock, scopes, error.errorCode, error.errorMessage);
            return reject(result);
        });
    });
}
exports.CheckUpnMatchIdToken = CheckUpnMatchIdToken;
/**
 * Log the telemetry data points in the provided or Otel pipeline.
 * @param result - the result to log the telemetry data points into
 * @param correlationId - the same correlation ID exists under the caller's context, the same correlation ID will passed on to MSAL.js for unified experience
 * @param applications - the application ID in the access token request
 * @param timerClock - the timerClock to log the time duration
 * @param errorCode - the error code included in the exception, if any
 * @param errorMessage - the error message included in the exception, if any
 */
function LogTelemetryDataFieldsForLoad(result, correlationId, applications, timerClock, errorCode, errorMessage) {
    if (!authConfig.telemetryInstance && typeof OTel === "undefined") {
        // For data fields that are null, blank or empty, the value is set to "unknown" at this point
        // based on office-online-otel documentation: https://office.visualstudio.com/OC/_git/office-online-ui?path=%2Fpackages%2Foffice-online-otel%2FREADME.md&version=GBmaster
        // Null or undefined value will cause exceptions down the line.
        var dataFields = [
            { name: Constants_1.Constants.Telemetry.Duration, int64: timerClock.ms },
            { name: Constants_1.Constants.Telemetry.Succeeded, bool: true },
            { name: Constants_1.Constants.Telemetry.IdentityProvider, string: authConfig.idp.toLowerCase() },
            { name: Constants_1.Constants.Telemetry.CorrelationId, string: correlationId ? correlationId : 'unknown' },
            { name: Constants_1.Constants.Telemetry.LoadedApplicationCount, int64: applications ? applications.length : 0 },
            { name: Constants_1.Constants.Telemetry.ErrorCode, string: errorCode ? errorCode : 'unknown' },
            { name: Constants_1.Constants.Telemetry.ErrorMessage, string: errorMessage ? errorMessage : 'unknown' },
        ];
        api_js_1.sendTelemetryEvent({
            name: Constants_1.Constants.Telemetry.LoadTelemetryName,
            dataFields: dataFields
        });
    }
    return {
        Telemetry: {
            timeToLoad: timerClock.ms,
            succeeded: result,
            idp: authConfig.idp.toLowerCase(),
            correlationId: correlationId ? correlationId : '',
            loadedApplicationCount: applications ? applications.length : 0,
            errorCode: errorCode ? errorCode : undefined,
            errorMessage: errorMessage ? errorMessage : undefined,
        }
    };
}
/**
 * Log the telemetry data points in the provided {@link IBrowserAuthResult} or Otel pipeline.
 * @param result - the provided {@link IBrowserAuthResult} to log the telemetry data points into
 * @param correlationId - the same correlation ID exists under the caller's context, the same correlation ID will passed on to MSAL.js for unified experience
 * @param applicationId - the application ID in the access token request
 * @param forThirdParty - If true, treats the caller as third-party and avoids sending PII telemetry
 * @param timerClock - the timerClock to log the time duration
 * @param scopes - the scopes in the acquire token request
 */
function LogTelemetryDataFieldsForGetToken(result, correlationId, applicationId, forThirdParty, timerClock, scopes) {
    if (!authConfig.telemetryInstance && typeof OTel === "undefined") {
        // For data fields that are null, blank or empty, the value is set to "unknown" at this point
        // based on office-online-otel documentation: https://office.visualstudio.com/OC/_git/office-online-ui?path=%2Fpackages%2Foffice-online-otel%2FREADME.md&version=GBmaster
        // Null or undefined value will cause exceptions down the line.
        var dataFields = [
            { name: Constants_1.Constants.Telemetry.Duration, int64: timerClock.ms },
            { name: Constants_1.Constants.Telemetry.Succeeded, string: result.Token ? false : true },
            { name: Constants_1.Constants.Telemetry.IdentityProvider, string: authConfig.idp.toLowerCase() },
            { name: Constants_1.Constants.Telemetry.ApplicationId, string: applicationId },
            { name: Constants_1.Constants.Telemetry.TokenScope, string: (scopes && !forThirdParty) ? scopes.toString() : 'unknown' },
            { name: Constants_1.Constants.Telemetry.CorrelationId, string: correlationId ? correlationId : 'unknown' },
            { name: Constants_1.Constants.Telemetry.ErrorCode, string: result.ErrorCode ? result.ErrorCode : 'unknown' },
            { name: Constants_1.Constants.Telemetry.ErrorMessage, string: (result.ErrorMessage && !forThirdParty) ? result.ErrorMessage : 'unknown' },
            { name: Constants_1.Constants.Telemetry.ErrorCodeForGetToken, string: result.ErrorCode ? result.ErrorCode : 'unknown' },
            { name: Constants_1.Constants.Telemetry.ErrorMessageForGetToken, string: (result.ErrorMessage && !forThirdParty) ? result.ErrorMessage : 'unknown' },
        ];
        api_js_1.sendTelemetryEvent({
            name: Constants_1.Constants.Telemetry.GetTokenTelemetryName,
            dataFields: dataFields
        });
    }
    result.Telemetry = {
        timeToGetToken: timerClock.ms,
        succeeded: result.Token ? false : true,
        idp: authConfig.idp.toLowerCase(),
        applicationId: applicationId,
        tokenScope: scopes && !forThirdParty ? scopes.toString() : undefined,
        correlationId: correlationId,
        errorCode: result.ErrorCode ? result.ErrorCode : undefined,
        errorMessage: result.ErrorMessage && !forThirdParty ? result.ErrorMessage : undefined,
        errorCodeForGetToken: result.ErrorCode ? result.ErrorCode : undefined,
        errorMessageForGetToken: result.ErrorMessage && !forThirdParty ? result.ErrorMessage : undefined,
        fromCache: result.MsalResult ? !!result.MsalResult.fromCache : false,
        expiresOn: result.MsalResult && result.MsalResult.expiresOn ? result.MsalResult.expiresOn.getTime() / 1000 : undefined
    };
    return result;
}
/**
 * * Log the telemetry data points in the provided {@link IBrowserAuthUpnResult} or Otel pipeline.
 * @param result - the provided {@link IBrowserAuthUpnResult} to log the telemetry data points into
 * @param correlationId - the same correlation ID exists under the caller's context, the same correlation ID will passed on to MSAL.js for unified experience
 * @param applicationId - the application ID in the access token request
 * @param timerClock - the timerClock to log the time duration
 * @param scopes - the scopes in the acquire token request
 * @param errorCode - the error code included in the exception, if any
 * @param errorMessage - the error message included in the exception, if any
 */
function LogTelemetryDataFieldsForCheckUpn(result, correlationId, applicationId, timerClock, scopes, errorCode, errorMessage) {
    // For data fields that are null, blank or empty, the value is set to "unknown" at this point
    // based on office-online-otel documentation: https://office.visualstudio.com/OC/_git/office-online-ui?path=%2Fpackages%2Foffice-online-otel%2FREADME.md&version=GBmaster
    // Null or undefined value will cause exceptions down the line.
    var dataFields = [
        { name: Constants_1.Constants.Telemetry.Duration, int64: timerClock.ms },
        { name: Constants_1.Constants.Telemetry.Succeeded, bool: errorCode ? false : true },
        { name: Constants_1.Constants.Telemetry.IdentityProvider, string: authConfig.idp.toLowerCase() },
        { name: Constants_1.Constants.Telemetry.ApplicationId, string: applicationId },
        { name: Constants_1.Constants.Telemetry.TokenScope, string: scopes ? scopes.toString() : 'unknown' },
        { name: Constants_1.Constants.Telemetry.CorrelationId, string: correlationId ? correlationId : 'unknown' },
        { name: Constants_1.Constants.Telemetry.ErrorCode, string: errorCode ? errorCode : 'unknown' },
        { name: Constants_1.Constants.Telemetry.ErrorMessage, string: errorMessage ? errorMessage : 'unknown' },
        { name: Constants_1.Constants.Telemetry.ErrorCodeForCheckUpn, string: errorCode ? errorCode : 'unknown' },
        { name: Constants_1.Constants.Telemetry.ErrorMessageForCheckUpn, string: errorMessage ? errorMessage : 'unknown' } //TODO: Deprecate after partner migration
    ];
    if (!authConfig.telemetryInstance && typeof OTel === "undefined") {
        api_js_1.sendTelemetryEvent({
            name: Constants_1.Constants.Telemetry.CheckUpnTelemetryName,
            dataFields: dataFields
        });
    }
    result.Telemetry = {
        timeToCheckUPN: timerClock.ms,
        succeeded: errorCode ? false : true,
        idp: authConfig.idp.toLowerCase(),
        applicationId: applicationId,
        tokenScope: scopes ? scopes.toString() : undefined,
        correlationId: correlationId,
        errorCode: errorCode,
        errorMessage: errorMessage,
        errorCodeForCheckUPN: errorCode,
        errorMessageForCheckUPN: errorMessage //TODO: Deprecate after partner migration
    };
}
/**
 * Logout current user from all registered applications
 */
function Logout() {
    applications.forEach(function (entry) {
        LoggingUtils_1.LoggingUtils.log("[Logout] application calls logout");
        entry.application.logout();
    });
    // Also clear registered applications
    exports.applications = applications = new Array();
}
exports.Logout = Logout;


/***/ }),

/***/ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/Constants.ts":
/*!**********************************************************************************!*\
  !*** ./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/Constants.ts ***!
  \**********************************************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var Constants;
(function (Constants) {
    var IdentityProvider = /** @class */ (function () {
        function IdentityProvider() {
        }
        /**
         * AAD
         */
        IdentityProvider.Aad = "aad";
        /**
         * MSA
         */
        IdentityProvider.Msa = "msa";
        return IdentityProvider;
    }());
    Constants.IdentityProvider = IdentityProvider;
    var PostMessageType = /** @class */ (function () {
        function PostMessageType() {
        }
        PostMessageType.RequestAuthConfig = "RequestAuthConfig";
        PostMessageType.ResponseAuthConfig = "ResponseAuthConfig";
        PostMessageType.iFrameIdTokenPrefix = "msalIdTokenFrame";
        PostMessageType.iFramePrefix = "msalRenewFrame";
        return PostMessageType;
    }());
    Constants.PostMessageType = PostMessageType;
    var Authority = /** @class */ (function () {
        function Authority() {
        }
        /**
         * Prod
         */
        Authority.Prod = "https://login.microsoftonline.com/";
        /**
         * Legacy Prod
         */
        Authority.ProdLegacy = "https://login.windows.net/";
        /**
         * PPE
         */
        Authority.Ppe = "https://login.windows-ppe.net/";
        Authority.AadSuffix = "organizations";
        Authority.MsaSuffix = "consumers";
        return Authority;
    }());
    Constants.Authority = Authority;
    var Telemetry = /** @class */ (function () {
        function Telemetry() {
        }
        Telemetry.OtelInstance = "otel";
        Telemetry.CheckUpnTelemetryName = "Office.Identity.WebAuth.BrowserAuth.CheckUpn";
        Telemetry.GetTokenTelemetryName = "Office.Identity.WebAuth.BrowserAuth.GetToken";
        Telemetry.LoadTelemetryName = "Office.Identity.WebAuth.BrowserAuth.Load";
        Telemetry.Duration = "Duration";
        Telemetry.Succeeded = "Succeeded";
        Telemetry.ApplicationId = "ApplicationId";
        Telemetry.CorrelationId = "CorrelationId";
        Telemetry.IdentityProvider = "IdentityProvider";
        Telemetry.LoadedApplicationCount = "LoadedApplicationCount";
        Telemetry.TokenScope = "TokenScope";
        Telemetry.ErrorCode = "ErrorCode";
        Telemetry.ErrorMessage = "ErrorMessage";
        Telemetry.ErrorCodeForLoad = "ErrorCodeForLoad";
        Telemetry.ErrorMessageForLoad = "ErrorMessageForLoad";
        Telemetry.ErrorCodeForGetToken = "ErrorCodeForGetToken";
        Telemetry.ErrorMessageForGetToken = "ErrorMessageForGetToken";
        Telemetry.ErrorCodeForCheckUpn = "ErrorCodeForCheckUpn";
        Telemetry.ErrorMessageForCheckUpn = "ErrorMessageForCheckUpn";
        return Telemetry;
    }());
    Constants.Telemetry = Telemetry;
})(Constants = exports.Constants || (exports.Constants = {}));


/***/ }),

/***/ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/ExtractUtils.ts":
/*!*************************************************************************************!*\
  !*** ./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/ExtractUtils.ts ***!
  \*************************************************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
/**
 * Internal Module contains utility methods for extracting tokens
 */
var LoggingUtils_1 = __webpack_require__(/*! ./LoggingUtils */ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/LoggingUtils.ts");
var ExtractUtils = /** @class */ (function () {
    function ExtractUtils() {
    }
    /**
     * Extract token by decoding the RAW token
     *
     * @param encodedToken
     */
    ExtractUtils.extractToken = function (encodedToken) {
        var decodedToken = this.decodeJwt(encodedToken);
        if (!decodedToken) {
            return null;
        }
        try {
            var base64Token = decodedToken.JWSPayload;
            var base64Decoded = this.base64Decode(base64Token);
            if (!base64Decoded) {
                LoggingUtils_1.LoggingUtils.log("The returned token could not be base64 url safe decoded.");
                return null;
            }
            // ECMA script has JSON built-in support
            return JSON.parse(base64Decoded);
        }
        catch (err) {
            LoggingUtils_1.LoggingUtils.error("The returned token could not be decoded" + err);
        }
        return null;
    };
    ;
    /**
     * Decodes a base64 encoded string.
     *
     * @param input
     */
    ExtractUtils.base64Decode = function (input) {
        var encodedString = input.replace(/-/g, "+").replace(/_/g, "/");
        switch (encodedString.length % 4) {
            case 0:
                break;
            case 2:
                encodedString += "==";
                break;
            case 3:
                encodedString += "=";
                break;
            default:
                throw new Error("Invalid base64 string");
        }
        return decodeURIComponent(atob(encodedString).split("").map(function (c) {
            return "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(""));
    };
    ;
    /**
     * decode a JWT
     *
     * @param jwtToken
     */
    ExtractUtils.decodeJwt = function (jwtToken) {
        if (jwtToken === "undefined" || !jwtToken || 0 === jwtToken.length) {
            return null;
        }
        var idTokenPartsRegex = /^([^\.\s]*)\.([^\.\s]+)\.([^\.\s]*)$/;
        var matches = idTokenPartsRegex.exec(jwtToken);
        if (!matches || matches.length < 4) {
            LoggingUtils_1.LoggingUtils.warn("The returned access_token is not parseable.");
            return null;
        }
        var crackedToken = {
            header: matches[1],
            JWSPayload: matches[2],
            JWSSig: matches[3]
        };
        return crackedToken;
    };
    ;
    return ExtractUtils;
}());
exports.ExtractUtils = ExtractUtils;


/***/ }),

/***/ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/LoggingUtils.ts":
/*!*************************************************************************************!*\
  !*** ./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/LoggingUtils.ts ***!
  \*************************************************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
/**
 * Internal Module contains utility methods for logging
 */
var LoggingUtils = /** @class */ (function () {
    function LoggingUtils() {
    }
    /**
     * Logs message to the console
     * @param message - message which was passed in
     * @param shouldLog - should this message be logged
     */
    LoggingUtils.log = function (message, shouldLog) {
        if (!this.shouldProceed(shouldLog)) {
            return false;
        }
        console.log(message);
        return true;
    };
    /**
     * Logs warning message to the console
     * @param message - message which was passed in
     * @param shouldLog - should this message be logged
     */
    LoggingUtils.warn = function (message, shouldLog) {
        if (!this.shouldProceed(shouldLog)) {
            return false;
        }
        console.warn(message);
        return true;
    };
    /**
     * Logs error message to the console
     * @param message - message which was passed in
     * @param shouldLog - should this message be logged
     */
    LoggingUtils.error = function (message, shouldLog) {
        if (!this.shouldProceed(shouldLog)) {
            return false;
        }
        console.error(message);
        return true;
    };
    /**
     * Returns if we should proceed with logging or not
     * @param shouldLog - should this message be logged
     */
    LoggingUtils.shouldProceed = function (shouldLog) {
        if (shouldLog != null && shouldLog !== undefined && !shouldLog) {
            return false;
        }
        if (shouldLog == null && shouldLog === undefined && !this.enableConsoleLogging) {
            return false;
        }
        return true;
    };
    LoggingUtils.enableConsoleLogging = false;
    return LoggingUtils;
}());
exports.LoggingUtils = LoggingUtils;


/***/ }),

/***/ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/TimerUtils.ts":
/*!***********************************************************************************!*\
  !*** ./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/TimerUtils.ts ***!
  \***********************************************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
/**
 * Module includes timer related utilities
 */
var TimerUtils = /** @class */ (function () {
    function TimerUtils() {
    }
    /**
     * Timer function
     */
    TimerUtils.timer = function () {
        var timeStart = new Date().getTime();
        return {
            /**
             * Returns time in seconds (example: 500)
             */
            get seconds() {
                var seconds = Math.ceil((new Date().getTime() - timeStart) / 1000);
                return seconds;
            },
            /**
             * Returns time in Milliseconds (example: 2000)
             */
            get ms() {
                var ms = (new Date().getTime() - timeStart);
                return ms;
            },
            /**
             * Returns formatted time in seconds (example: 500s)
             */
            get formattedSeconds() {
                var seconds = Math.ceil(this.seconds / 1000) + "s";
                return seconds;
            },
            /**
             * Returns formatted time in Milliseconds (example: 2000ms)
             */
            get formattedMs() {
                var ms = this.ms + "ms";
                return ms;
            }
        };
    };
    return TimerUtils;
}());
exports.TimerUtils = TimerUtils;


/***/ }),

/***/ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/ValidateUtils.ts":
/*!**************************************************************************************!*\
  !*** ./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/ValidateUtils.ts ***!
  \**************************************************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/**
 * Internal Module contains utility methods for validation
 */
Object.defineProperty(exports, "__esModule", { value: true });
var ExtractUtils_1 = __webpack_require__(/*! ./ExtractUtils */ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/ExtractUtils.ts");
var LoggingUtils_1 = __webpack_require__(/*! ./LoggingUtils */ "./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/Utils/LoggingUtils.ts");
var ValidateUtils = /** @class */ (function () {
    function ValidateUtils() {
    }
    /**
     * Construct the scope for V2 endpoint calls.
     * @param target - resource for V1 token, scope for V2 token
     * @returns the right scope to make calls to V2 endpoint.
     */
    ValidateUtils.getScope = function (target) {
        // To consume V2 endpoint, "/.default" needs to be added for given resources.
        var resourcePrefix = ["HTTPS:", "API:"];
        if (resourcePrefix.some(function (prefix) { return target.toLocaleUpperCase().startsWith(prefix); }) || this.isGuid(target)) {
            return target + "/.default";
        }
        // Other cases could be that it is acquiring V2 tokens with scopes "ConnectedServices.ReadWrite" etc
        // or wl.skydrive
        return target;
    };
    /**
     * Check whether the given string is in GUID format or not.
     * @param str - provided string for format checking.
     * @returns true if the string is in GUID format, returns false otherwise.
     */
    ValidateUtils.isGuid = function (str) {
        // Checking GUID based on the GUID format
        var regexGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
        return regexGuid.test(str);
    };
    /**
     * Verify the upn in the config matches the upn in the id token
     * @param token - the token to extract the upn from.
     * @returns true if there is a match or there is no upn in the config
     */
    ValidateUtils.upnMatchesUpnFromIdToken = function (token, authConfig) {
        if (!this.enableUpnCheck) {
            // Skip upn check if enableUpnCheck is not true
            LoggingUtils_1.LoggingUtils.log('Skip Upn check of id token, returning true');
            return true;
        }
        if (!authConfig || !authConfig.upn) {
            LoggingUtils_1.LoggingUtils.log('Upn does not exist in the configuration, returning true');
            return true;
        }
        var idToken = ExtractUtils_1.ExtractUtils.extractToken(token);
        // IdToken extraction would not work for future encrypted JWE tokens,
        // If cannot be extracted, also return true
        if (!idToken || (idToken && idToken.preferred_username && idToken.preferred_username.toLowerCase() === authConfig.upn.toLowerCase())) {
            return true;
        }
        LoggingUtils_1.LoggingUtils.log('provided Upn does not match Upn extracted from id token');
        return false;
    };
    /**
     * Verify the upn in the id token matches the Upn in the access token
     * @param token - the token to extract the upn from
     * @returns true if there is a match or there is no upn in the id token or access token
     */
    ValidateUtils.upnMatchesUpnFromAccessToken = function (token, authConfig) {
        if (!this.enableUpnCheck) {
            // Skip upn check if enableUpnCheck is not true
            LoggingUtils_1.LoggingUtils.log('Skip Upn check of access token, returning true');
            return true;
        }
        if (!authConfig || !authConfig.upn) {
            LoggingUtils_1.LoggingUtils.log('Upn does not exist in the configuration, returning true');
            return true;
        }
        var accessToken = ExtractUtils_1.ExtractUtils.extractToken(token);
        // AccessToken extraction would not work for future encrypted JWE tokens,
        // If cannot be extracted, also return true
        if (!accessToken || (accessToken && accessToken.upn && accessToken.upn.toLowerCase() === authConfig.upn.toLowerCase())) {
            return true;
        }
        LoggingUtils_1.LoggingUtils.log('provided Upn does not match Upn extracted from access token');
        return false;
    };
    ValidateUtils.enableUpnCheck = false;
    return ValidateUtils;
}());
exports.ValidateUtils = ValidateUtils;


/***/ }),

/***/ 0:
/*!***********************************************************************************************************************************************!*\
  !*** multi ./packages/Microsoft.Office.WebAuth.BrowserAuth/lib/api.js ./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/BrowserAuth.ts ***!
  \***********************************************************************************************************************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

__webpack_require__(/*! D:\Repos\WebAuthJS\src\packages\Microsoft.Office.WebAuth.BrowserAuth\lib\api.js */"./packages/Microsoft.Office.WebAuth.BrowserAuth/lib/api.js");
module.exports = __webpack_require__(/*! D:\Repos\WebAuthJS\src\packages\Microsoft.Office.WebAuth.BrowserAuth\scripts\BrowserAuth.ts */"./packages/Microsoft.Office.WebAuth.BrowserAuth/scripts/BrowserAuth.ts");


/***/ })

/******/ });
//# sourceMappingURL=webauth.browserauth.js.map
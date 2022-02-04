var OSFPerformance;
(function (OSFPerformance) {
    OSFPerformance.officeExecuteStartDate = 0;
    OSFPerformance.officeExecuteStart = 0;
    OSFPerformance.officeExecuteEnd = 0;
    OSFPerformance.hostInitializationStart = 0;
    OSFPerformance.hostInitializationEnd = 0;
    OSFPerformance.totalJSHeapSize = 0;
    OSFPerformance.usedJSHeapSize = 0;
    OSFPerformance.jsHeapSizeLimit = 0;
    OSFPerformance.getAppContextStart = 0;
    OSFPerformance.getAppContextEnd = 0;
    OSFPerformance.createOMEnd = 0;
    OSFPerformance.officeOnReady = 0;
    OSFPerformance.hostSpecificFileName = "";
    function now() {
        if (performance && performance.now) {
            return performance.now();
        }
        else {
            return 0;
        }
    }
    OSFPerformance.now = now;
    function getTotalJSHeapSize() {
        if (typeof (performance) !== 'undefined' && performance.memory) {
            return performance.memory.totalJSHeapSize;
        }
        else {
            return 0;
        }
    }
    OSFPerformance.getTotalJSHeapSize = getTotalJSHeapSize;
    function getUsedJSHeapSize() {
        if (typeof (performance) !== 'undefined' && performance.memory) {
            return performance.memory.usedJSHeapSize;
        }
        else {
            return 0;
        }
    }
    OSFPerformance.getUsedJSHeapSize = getUsedJSHeapSize;
    function getJSHeapSizeLimit() {
        if (typeof (performance) !== 'undefined' && performance.memory) {
            return performance.memory.jsHeapSizeLimit;
        }
        else {
            return 0;
        }
    }
    OSFPerformance.getJSHeapSizeLimit = getJSHeapSizeLimit;
})(OSFPerformance || (OSFPerformance = {}));
;
OSFPerformance.officeExecuteStartDate = Date.now();
OSFPerformance.officeExecuteStart = OSFPerformance.now();



/* Office JavaScript API library */

/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/


/*
    Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.

    This file also contains the following Promise implementation (with a few small modifications):
        * @overview es6-promise - a tiny implementation of Promises/A+.
        * @copyright Copyright (c) 2014 Yehuda Katz, Tom Dale, Stefan Penner and contributors (Conversion to ES6 API by Jake Archibald)
        * @license   Licensed under MIT license
        *            See https://raw.githubusercontent.com/jakearchibald/es6-promise/master/LICENSE
        * @version   2.3.0
*/
var OSF = OSF || {};
OSF.HostSpecificFileVersionDefault = "16.00";
OSF.HostSpecificFileVersionMap = {
    "access": {
        "web": "16.00"
    },
    "agavito": {
        "winrt": "16.00"
    },
    "excel": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    },
    "onenote": {
        "android": "16.00",
        "web": "16.00",
        "win32": "16.00",
        "winrt": "16.00"
    },
    "outlook": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.01",
        "win32": "16.02"
    },
    "powerpoint": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    },
    "project": {
        "win32": "16.00"
    },
    "sway": {
        "web": "16.00"
    },
    "word": {
        "ios": "16.00",
        "mac": "16.00",
        "web": "16.00",
        "win32": "16.01",
        "winrt": "16.00"
    },
    "visio": {
        "web": "16.00",
        "win32": "16.00"
    }
};
OSF.SupportedLocales = {
    "ar-sa": true,
    "bg-bg": true,
    "bn-in": true,
    "ca-es": true,
    "cs-cz": true,
    "da-dk": true,
    "de-de": true,
    "el-gr": true,
    "en-us": true,
    "es-es": true,
    "et-ee": true,
    "eu-es": true,
    "fa-ir": true,
    "fi-fi": true,
    "fr-fr": true,
    "gl-es": true,
    "he-il": true,
    "hi-in": true,
    "hr-hr": true,
    "hu-hu": true,
    "id-id": true,
    "it-it": true,
    "ja-jp": true,
    "kk-kz": true,
    "ko-kr": true,
    "lo-la": true,
    "lt-lt": true,
    "lv-lv": true,
    "ms-my": true,
    "nb-no": true,
    "nl-nl": true,
    "nn-no": true,
    "pl-pl": true,
    "pt-br": true,
    "pt-pt": true,
    "ro-ro": true,
    "ru-ru": true,
    "sk-sk": true,
    "sl-si": true,
    "sr-cyrl-cs": true,
    "sr-cyrl-rs": true,
    "sr-latn-cs": true,
    "sr-latn-rs": true,
    "sv-se": true,
    "th-th": true,
    "tr-tr": true,
    "uk-ua": true,
    "ur-pk": true,
    "vi-vn": true,
    "zh-cn": true,
    "zh-tw": true
};
OSF.AssociatedLocales = {
    ar: "ar-sa",
    bg: "bg-bg",
    bn: "bn-in",
    ca: "ca-es",
    cs: "cs-cz",
    da: "da-dk",
    de: "de-de",
    el: "el-gr",
    en: "en-us",
    es: "es-es",
    et: "et-ee",
    eu: "eu-es",
    fa: "fa-ir",
    fi: "fi-fi",
    fr: "fr-fr",
    gl: "gl-es",
    he: "he-il",
    hi: "hi-in",
    hr: "hr-hr",
    hu: "hu-hu",
    id: "id-id",
    it: "it-it",
    ja: "ja-jp",
    kk: "kk-kz",
    ko: "ko-kr",
    lo: "lo-la",
    lt: "lt-lt",
    lv: "lv-lv",
    ms: "ms-my",
    nb: "nb-no",
    nl: "nl-nl",
    nn: "nn-no",
    pl: "pl-pl",
    pt: "pt-br",
    ro: "ro-ro",
    ru: "ru-ru",
    sk: "sk-sk",
    sl: "sl-si",
    sr: "sr-cyrl-cs",
    sv: "sv-se",
    th: "th-th",
    tr: "tr-tr",
    uk: "uk-ua",
    ur: "ur-pk",
    vi: "vi-vn",
    zh: "zh-cn"
};
OSF.getSupportedLocale = function OSF$getSupportedLocale(locale, defaultLocale) {
    if (defaultLocale === void 0) { defaultLocale = "en-us"; }
    if (!locale) {
        return defaultLocale;
    }
    var supportedLocale;
    locale = locale.toLowerCase();
    if (locale in OSF.SupportedLocales) {
        supportedLocale = locale;
    }
    else {
        var localeParts = locale.split('-', 1);
        if (localeParts && localeParts.length > 0) {
            supportedLocale = OSF.AssociatedLocales[localeParts[0]];
        }
    }
    if (!supportedLocale) {
        supportedLocale = defaultLocale;
    }
    return supportedLocale;
};
var ScriptLoading;
(function (ScriptLoading) {
    var ScriptInfo = (function () {
        function ScriptInfo(url, isReady, hasStarted, timer, pendingCallback) {
            this.url = url;
            this.isReady = isReady;
            this.hasStarted = hasStarted;
            this.timer = timer;
            this.hasError = false;
            this.pendingCallbacks = [];
            this.pendingCallbacks.push(pendingCallback);
        }
        return ScriptInfo;
    }());
    var ScriptTelemetry = (function () {
        function ScriptTelemetry(scriptId, startTime, msResponseTime) {
            this.scriptId = scriptId;
            this.startTime = startTime;
            this.msResponseTime = msResponseTime;
        }
        return ScriptTelemetry;
    }());
    var LoadScriptHelper = (function () {
        function LoadScriptHelper(constantNames) {
            if (constantNames === void 0) { constantNames = {
                OfficeJS: "office.js",
                OfficeDebugJS: "office.debug.js"
            }; }
            this.constantNames = constantNames;
            this.defaultScriptLoadingTimeout = 10000;
            this.loadedScriptByIds = {};
            this.scriptTelemetryBuffer = [];
            this.osfControlAppCorrelationId = "";
            this.basePath = null;
        }
        LoadScriptHelper.prototype.isScriptLoading = function (id) {
            return !!(this.loadedScriptByIds[id] && this.loadedScriptByIds[id].hasStarted);
        };
        LoadScriptHelper.prototype.getOfficeJsBasePath = function () {
            if (this.basePath) {
                return this.basePath;
            }
            else {
                var getScriptBase = function (scriptSrc, scriptNameToCheck) {
                    var scriptBase, indexOfJS, scriptSrcLowerCase;
                    scriptSrcLowerCase = scriptSrc.toLowerCase();
                    indexOfJS = scriptSrcLowerCase.indexOf(scriptNameToCheck);
                    if (indexOfJS >= 0 && indexOfJS === (scriptSrc.length - scriptNameToCheck.length) && (indexOfJS === 0 || scriptSrc.charAt(indexOfJS - 1) === '/' || scriptSrc.charAt(indexOfJS - 1) === '\\')) {
                        scriptBase = scriptSrc.substring(0, indexOfJS);
                    }
                    else if (indexOfJS >= 0
                        && indexOfJS < (scriptSrc.length - scriptNameToCheck.length)
                        && scriptSrc.charAt(indexOfJS + scriptNameToCheck.length) === '?'
                        && (indexOfJS === 0 || scriptSrc.charAt(indexOfJS - 1) === '/' || scriptSrc.charAt(indexOfJS - 1) === '\\')) {
                        scriptBase = scriptSrc.substring(0, indexOfJS);
                    }
                    return scriptBase;
                };
                var scripts = document.getElementsByTagName("script");
                var scriptsCount = scripts.length;
                var officeScripts = [this.constantNames.OfficeJS, this.constantNames.OfficeDebugJS];
                var officeScriptsCount = officeScripts.length;
                var i, j;
                for (i = 0; !this.basePath && i < scriptsCount; i++) {
                    if (scripts[i].src) {
                        for (j = 0; !this.basePath && j < officeScriptsCount; j++) {
                            this.basePath = getScriptBase(scripts[i].src, officeScripts[j]);
                        }
                    }
                }
                return this.basePath;
            }
        };
        LoadScriptHelper.prototype.loadScript = function (url, scriptId, callback, highPriority, timeoutInMs) {
            this.loadScriptInternal(url, scriptId, callback, highPriority, timeoutInMs);
        };
        LoadScriptHelper.prototype.loadScriptParallel = function (url, scriptId, timeoutInMs) {
            this.loadScriptInternal(url, scriptId, null, false, timeoutInMs);
        };
        LoadScriptHelper.prototype.waitForFunction = function (scriptLoadTest, callback, numberOfTries, delay) {
            var attemptsRemaining = numberOfTries;
            var timerId;
            var validateFunction = function () {
                attemptsRemaining--;
                if (scriptLoadTest()) {
                    callback(true);
                    return;
                }
                else if (attemptsRemaining > 0) {
                    timerId = window.setTimeout(validateFunction, delay);
                    attemptsRemaining--;
                }
                else {
                    window.clearTimeout(timerId);
                    callback(false);
                }
            };
            validateFunction();
        };
        LoadScriptHelper.prototype.waitForScripts = function (ids, callback) {
            var _this = this;
            if (this.invokeCallbackIfScriptsReady(ids, callback) == false) {
                for (var i = 0; i < ids.length; i++) {
                    var id = ids[i];
                    var loadedScriptEntry = this.loadedScriptByIds[id];
                    if (loadedScriptEntry) {
                        loadedScriptEntry.pendingCallbacks.push(function () {
                            _this.invokeCallbackIfScriptsReady(ids, callback);
                        });
                    }
                }
            }
        };
        LoadScriptHelper.prototype.logScriptLoading = function (scriptId, startTime, msResponseTime) {
            startTime = Math.floor(startTime);
            if (OSF.AppTelemetry && OSF.AppTelemetry.onScriptDone) {
                if (OSF.AppTelemetry.onScriptDone.length == 3) {
                    OSF.AppTelemetry.onScriptDone(scriptId, startTime, msResponseTime);
                }
                else {
                    OSF.AppTelemetry.onScriptDone(scriptId, startTime, msResponseTime, this.osfControlAppCorrelationId);
                }
            }
            else {
                var scriptTelemetry = new ScriptTelemetry(scriptId, startTime, msResponseTime);
                this.scriptTelemetryBuffer.push(scriptTelemetry);
            }
        };
        LoadScriptHelper.prototype.setAppCorrelationId = function (appCorrelationId) {
            this.osfControlAppCorrelationId = appCorrelationId;
        };
        LoadScriptHelper.prototype.invokeCallbackIfScriptsReady = function (ids, callback) {
            var hasError = false;
            for (var i = 0; i < ids.length; i++) {
                var id = ids[i];
                var loadedScriptEntry = this.loadedScriptByIds[id];
                if (!loadedScriptEntry) {
                    loadedScriptEntry = new ScriptInfo("", false, false, null, null);
                    this.loadedScriptByIds[id] = loadedScriptEntry;
                }
                if (loadedScriptEntry.isReady == false) {
                    return false;
                }
                else if (loadedScriptEntry.hasError) {
                    hasError = true;
                }
            }
            callback(!hasError);
            return true;
        };
        LoadScriptHelper.prototype.getScriptEntryByUrl = function (url) {
            for (var key in this.loadedScriptByIds) {
                var scriptEntry = this.loadedScriptByIds[key];
                if (this.loadedScriptByIds.hasOwnProperty(key) && scriptEntry.url === url) {
                    return scriptEntry;
                }
            }
            return null;
        };
        LoadScriptHelper.prototype.loadScriptInternal = function (url, scriptId, callback, highPriority, timeoutInMs) {
            if (url) {
                var self = this;
                var doc = window.document;
                var loadedScriptEntry = (scriptId && this.loadedScriptByIds[scriptId]) ? this.loadedScriptByIds[scriptId] : this.getScriptEntryByUrl(url);
                if (!loadedScriptEntry || loadedScriptEntry.hasError || loadedScriptEntry.url.toLowerCase() != url.toLowerCase()) {
                    var script = doc.createElement("script");
                    script.type = "text/javascript";
                    if (scriptId) {
                        script.id = scriptId;
                    }
                    if (!loadedScriptEntry) {
                        loadedScriptEntry = new ScriptInfo(url, false, false, null, null);
                        this.loadedScriptByIds[(scriptId ? scriptId : url)] = loadedScriptEntry;
                    }
                    else {
                        loadedScriptEntry.url = url;
                        loadedScriptEntry.hasError = false;
                        loadedScriptEntry.isReady = false;
                    }
                    if (callback) {
                        if (highPriority) {
                            loadedScriptEntry.pendingCallbacks.unshift(callback);
                        }
                        else {
                            loadedScriptEntry.pendingCallbacks.push(callback);
                        }
                    }
                    var timeFromPageInit = -1;
                    if (window.performance && window.performance.now) {
                        timeFromPageInit = window.performance.now();
                    }
                    var startTime = (new Date()).getTime();
                    var logTelemetry = function (succeeded) {
                        if (scriptId) {
                            var totalTime = (new Date()).getTime() - startTime;
                            if (!succeeded) {
                                totalTime = -totalTime;
                            }
                            self.logScriptLoading(scriptId, timeFromPageInit, totalTime);
                        }
                        self.flushTelemetryBuffer();
                    };
                    var onLoadCallback = function OSF_OUtil_loadScript$onLoadCallback() {
                        if (!OSF._OfficeAppFactory.getLoggingAllowed() && (typeof OSF.AppTelemetry !== 'undefined')) {
                            OSF.AppTelemetry.enableTelemetry = false;
                        }
                        logTelemetry(true);
                        loadedScriptEntry.isReady = true;
                        if (loadedScriptEntry.timer != null) {
                            clearTimeout(loadedScriptEntry.timer);
                            delete loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = loadedScriptEntry.pendingCallbacks.shift();
                            if (currentCallback) {
                                var result = currentCallback(true);
                                if (result === false) {
                                    break;
                                }
                            }
                        }
                    };
                    var onLoadError = function () {
                        logTelemetry(false);
                        loadedScriptEntry.hasError = true;
                        loadedScriptEntry.isReady = true;
                        if (loadedScriptEntry.timer != null) {
                            clearTimeout(loadedScriptEntry.timer);
                            delete loadedScriptEntry.timer;
                        }
                        var pendingCallbackCount = loadedScriptEntry.pendingCallbacks.length;
                        for (var i = 0; i < pendingCallbackCount; i++) {
                            var currentCallback = loadedScriptEntry.pendingCallbacks.shift();
                            if (currentCallback) {
                                var result = currentCallback(false);
                                if (result === false) {
                                    break;
                                }
                            }
                        }
                    };
                    if (script.readyState) {
                        script.onreadystatechange = function () {
                            if (script.readyState == "loaded" || script.readyState == "complete") {
                                script.onreadystatechange = null;
                                onLoadCallback();
                            }
                        };
                    }
                    else {
                        script.onload = onLoadCallback;
                    }
                    script.onerror = onLoadError;
                    timeoutInMs = timeoutInMs || this.defaultScriptLoadingTimeout;
                    loadedScriptEntry.timer = setTimeout(onLoadError, timeoutInMs);
                    loadedScriptEntry.hasStarted = true;
                    script.setAttribute("crossOrigin", "anonymous");
                    script.src = url;
                    doc.getElementsByTagName("head")[0].appendChild(script);
                }
                else if (loadedScriptEntry.isReady) {
                    callback(true);
                }
                else {
                    if (highPriority) {
                        loadedScriptEntry.pendingCallbacks.unshift(callback);
                    }
                    else {
                        loadedScriptEntry.pendingCallbacks.push(callback);
                    }
                }
            }
        };
        LoadScriptHelper.prototype.flushTelemetryBuffer = function () {
            if (OSF.AppTelemetry && OSF.AppTelemetry.onScriptDone) {
                for (var i = 0; i < this.scriptTelemetryBuffer.length; i++) {
                    var scriptTelemetry = this.scriptTelemetryBuffer[i];
                    if (OSF.AppTelemetry.onScriptDone.length == 3) {
                        OSF.AppTelemetry.onScriptDone(scriptTelemetry.scriptId, scriptTelemetry.startTime, scriptTelemetry.msResponseTime);
                    }
                    else {
                        OSF.AppTelemetry.onScriptDone(scriptTelemetry.scriptId, scriptTelemetry.startTime, scriptTelemetry.msResponseTime, this.osfControlAppCorrelationId);
                    }
                }
                this.scriptTelemetryBuffer = [];
            }
        };
        return LoadScriptHelper;
    }());
    ScriptLoading.LoadScriptHelper = LoadScriptHelper;
})(ScriptLoading || (ScriptLoading = {}));
var OfficeExt;
(function (OfficeExt) {
    var HostName;
    (function (HostName) {
        var Host = (function () {
            function Host() {
                this.getDiagnostics = function _getDiagnostics(version) {
                    var diagnostics = {
                        host: this.getHost(),
                        version: (version || this.getDefaultVersion()),
                        platform: this.getPlatform()
                    };
                    return diagnostics;
                };
                this.platformRemappings = {
                    web: Microsoft.Office.WebExtension.PlatformType.OfficeOnline,
                    winrt: Microsoft.Office.WebExtension.PlatformType.Universal,
                    win32: Microsoft.Office.WebExtension.PlatformType.PC,
                    mac: Microsoft.Office.WebExtension.PlatformType.Mac,
                    ios: Microsoft.Office.WebExtension.PlatformType.iOS,
                    android: Microsoft.Office.WebExtension.PlatformType.Android
                };
                this.camelCaseMappings = {
                    powerpoint: Microsoft.Office.WebExtension.HostType.PowerPoint,
                    onenote: Microsoft.Office.WebExtension.HostType.OneNote
                };
                this.hostInfo = OSF._OfficeAppFactory.getHostInfo();
                this.getHost = this.getHost.bind(this);
                this.getPlatform = this.getPlatform.bind(this);
                this.getDiagnostics = this.getDiagnostics.bind(this);
            }
            Host.prototype.capitalizeFirstLetter = function (input) {
                if (input) {
                    return (input[0].toUpperCase() + input.slice(1).toLowerCase());
                }
                return input;
            };
            Host.getInstance = function () {
                if (Host.hostObj === undefined) {
                    Host.hostObj = new Host();
                }
                return Host.hostObj;
            };
            Host.prototype.getPlatform = function (appNumber) {
                if (this.hostInfo.hostPlatform) {
                    var hostPlatform = this.hostInfo.hostPlatform.toLowerCase();
                    if (this.platformRemappings[hostPlatform]) {
                        return this.platformRemappings[hostPlatform];
                    }
                }
                return null;
            };
            Host.prototype.getHost = function (appNumber) {
                if (this.hostInfo.hostType) {
                    var hostType = this.hostInfo.hostType.toLowerCase();
                    if (this.camelCaseMappings[hostType]) {
                        return this.camelCaseMappings[hostType];
                    }
                    hostType = this.capitalizeFirstLetter(this.hostInfo.hostType);
                    if (Microsoft.Office.WebExtension.HostType[hostType]) {
                        return Microsoft.Office.WebExtension.HostType[hostType];
                    }
                }
                return null;
            };
            Host.prototype.getDefaultVersion = function () {
                if (this.getHost()) {
                    return "16.0.0000.0000";
                }
                return null;
            };
            return Host;
        }());
        HostName.Host = Host;
    })(HostName = OfficeExt.HostName || (OfficeExt.HostName = {}));
})(OfficeExt || (OfficeExt = {}));
var Office;
(function (Office) {
    var _Internal;
    (function (_Internal) {
        var PromiseImpl;
        (function (PromiseImpl) {
            function Init() {
                return (function () {
                    "use strict";
                    function lib$es6$promise$utils$$objectOrFunction(x) {
                        return typeof x === 'function' || (typeof x === 'object' && x !== null);
                    }
                    function lib$es6$promise$utils$$isFunction(x) {
                        return typeof x === 'function';
                    }
                    function lib$es6$promise$utils$$isMaybeThenable(x) {
                        return typeof x === 'object' && x !== null;
                    }
                    var lib$es6$promise$utils$$_isArray;
                    if (!Array.isArray) {
                        lib$es6$promise$utils$$_isArray = function (x) {
                            return Object.prototype.toString.call(x) === '[object Array]';
                        };
                    }
                    else {
                        lib$es6$promise$utils$$_isArray = Array.isArray;
                    }
                    var lib$es6$promise$utils$$isArray = lib$es6$promise$utils$$_isArray;
                    var lib$es6$promise$asap$$len = 0;
                    var lib$es6$promise$asap$$toString = {}.toString;
                    var lib$es6$promise$asap$$vertxNext;
                    var lib$es6$promise$asap$$customSchedulerFn;
                    var lib$es6$promise$asap$$asap = function asap(callback, arg) {
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len] = callback;
                        lib$es6$promise$asap$$queue[lib$es6$promise$asap$$len + 1] = arg;
                        lib$es6$promise$asap$$len += 2;
                        if (lib$es6$promise$asap$$len === 2) {
                            if (lib$es6$promise$asap$$customSchedulerFn) {
                                lib$es6$promise$asap$$customSchedulerFn(lib$es6$promise$asap$$flush);
                            }
                            else {
                                lib$es6$promise$asap$$scheduleFlush();
                            }
                        }
                    };
                    function lib$es6$promise$asap$$setScheduler(scheduleFn) {
                        lib$es6$promise$asap$$customSchedulerFn = scheduleFn;
                    }
                    function lib$es6$promise$asap$$setAsap(asapFn) {
                        lib$es6$promise$asap$$asap = asapFn;
                    }
                    var lib$es6$promise$asap$$browserWindow = (typeof window !== 'undefined') ? window : undefined;
                    var lib$es6$promise$asap$$browserGlobal = lib$es6$promise$asap$$browserWindow || {};
                    var lib$es6$promise$asap$$BrowserMutationObserver = lib$es6$promise$asap$$browserGlobal.MutationObserver || lib$es6$promise$asap$$browserGlobal.WebKitMutationObserver;
                    var lib$es6$promise$asap$$isNode = typeof process !== 'undefined' && {}.toString.call(process) === '[object process]';
                    var lib$es6$promise$asap$$isWorker = typeof Uint8ClampedArray !== 'undefined' &&
                        typeof importScripts !== 'undefined' &&
                        typeof MessageChannel !== 'undefined';
                    function lib$es6$promise$asap$$useNextTick() {
                        var nextTick = process.nextTick;
                        var version = process.versions.node.match(/^(?:(\d+)\.)?(?:(\d+)\.)?(\*|\d+)$/);
                        if (Array.isArray(version) && version[1] === '0' && version[2] === '10') {
                            nextTick = setImmediate;
                        }
                        return function () {
                            nextTick(lib$es6$promise$asap$$flush);
                        };
                    }
                    function lib$es6$promise$asap$$useVertxTimer() {
                        return function () {
                            lib$es6$promise$asap$$vertxNext(lib$es6$promise$asap$$flush);
                        };
                    }
                    function lib$es6$promise$asap$$useMutationObserver() {
                        var iterations = 0;
                        var observer = new lib$es6$promise$asap$$BrowserMutationObserver(lib$es6$promise$asap$$flush);
                        var node = document.createTextNode('');
                        observer.observe(node, { characterData: true });
                        return function () {
                            node.data = (iterations = ++iterations % 2);
                        };
                    }
                    function lib$es6$promise$asap$$useMessageChannel() {
                        var channel = new MessageChannel();
                        channel.port1.onmessage = lib$es6$promise$asap$$flush;
                        return function () {
                            channel.port2.postMessage(0);
                        };
                    }
                    function lib$es6$promise$asap$$useSetTimeout() {
                        return function () {
                            setTimeout(lib$es6$promise$asap$$flush, 1);
                        };
                    }
                    var lib$es6$promise$asap$$queue = new Array(1000);
                    function lib$es6$promise$asap$$flush() {
                        for (var i = 0; i < lib$es6$promise$asap$$len; i += 2) {
                            var callback = lib$es6$promise$asap$$queue[i];
                            var arg = lib$es6$promise$asap$$queue[i + 1];
                            callback(arg);
                            lib$es6$promise$asap$$queue[i] = undefined;
                            lib$es6$promise$asap$$queue[i + 1] = undefined;
                        }
                        lib$es6$promise$asap$$len = 0;
                    }
                    var lib$es6$promise$asap$$scheduleFlush;
                    if (lib$es6$promise$asap$$isNode) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useNextTick();
                    }
                    else if (lib$es6$promise$asap$$isWorker) {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useMessageChannel();
                    }
                    else {
                        lib$es6$promise$asap$$scheduleFlush = lib$es6$promise$asap$$useSetTimeout();
                    }
                    function lib$es6$promise$$internal$$noop() { }
                    var lib$es6$promise$$internal$$PENDING = void 0;
                    var lib$es6$promise$$internal$$FULFILLED = 1;
                    var lib$es6$promise$$internal$$REJECTED = 2;
                    var lib$es6$promise$$internal$$GET_THEN_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                    function lib$es6$promise$$internal$$selfFullfillment() {
                        return new TypeError("You cannot resolve a promise with itself");
                    }
                    function lib$es6$promise$$internal$$cannotReturnOwn() {
                        return new TypeError('A promises callback cannot return that same promise.');
                    }
                    function lib$es6$promise$$internal$$getThen(promise) {
                        try {
                            return promise.then;
                        }
                        catch (error) {
                            lib$es6$promise$$internal$$GET_THEN_ERROR.error = error;
                            return lib$es6$promise$$internal$$GET_THEN_ERROR;
                        }
                    }
                    function lib$es6$promise$$internal$$tryThen(then, value, fulfillmentHandler, rejectionHandler) {
                        try {
                            then.call(value, fulfillmentHandler, rejectionHandler);
                        }
                        catch (e) {
                            return e;
                        }
                    }
                    function lib$es6$promise$$internal$$handleForeignThenable(promise, thenable, then) {
                        lib$es6$promise$asap$$asap(function (promise) {
                            var sealed = false;
                            var error = lib$es6$promise$$internal$$tryThen(then, thenable, function (value) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                if (thenable !== value) {
                                    lib$es6$promise$$internal$$resolve(promise, value);
                                }
                                else {
                                    lib$es6$promise$$internal$$fulfill(promise, value);
                                }
                            }, function (reason) {
                                if (sealed) {
                                    return;
                                }
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, reason);
                            }, 'Settle: ' + (promise._label || ' unknown promise'));
                            if (!sealed && error) {
                                sealed = true;
                                lib$es6$promise$$internal$$reject(promise, error);
                            }
                        }, promise);
                    }
                    function lib$es6$promise$$internal$$handleOwnThenable(promise, thenable) {
                        if (thenable._state === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, thenable._result);
                        }
                        else if (thenable._state === lib$es6$promise$$internal$$REJECTED) {
                            lib$es6$promise$$internal$$reject(promise, thenable._result);
                        }
                        else {
                            lib$es6$promise$$internal$$subscribe(thenable, undefined, function (value) {
                                lib$es6$promise$$internal$$resolve(promise, value);
                            }, function (reason) {
                                lib$es6$promise$$internal$$reject(promise, reason);
                            });
                        }
                    }
                    function lib$es6$promise$$internal$$handleMaybeThenable(promise, maybeThenable) {
                        if (maybeThenable.constructor === promise.constructor) {
                            lib$es6$promise$$internal$$handleOwnThenable(promise, maybeThenable);
                        }
                        else {
                            var then = lib$es6$promise$$internal$$getThen(maybeThenable);
                            if (then === lib$es6$promise$$internal$$GET_THEN_ERROR) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$GET_THEN_ERROR.error);
                            }
                            else if (then === undefined) {
                                lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                            }
                            else if (lib$es6$promise$utils$$isFunction(then)) {
                                lib$es6$promise$$internal$$handleForeignThenable(promise, maybeThenable, then);
                            }
                            else {
                                lib$es6$promise$$internal$$fulfill(promise, maybeThenable);
                            }
                        }
                    }
                    function lib$es6$promise$$internal$$resolve(promise, value) {
                        if (promise === value) {
                            lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$selfFullfillment());
                        }
                        else if (lib$es6$promise$utils$$objectOrFunction(value)) {
                            lib$es6$promise$$internal$$handleMaybeThenable(promise, value);
                        }
                        else {
                            lib$es6$promise$$internal$$fulfill(promise, value);
                        }
                    }
                    function lib$es6$promise$$internal$$publishRejection(promise) {
                        if (promise._onerror) {
                            promise._onerror(promise._result);
                        }
                        lib$es6$promise$$internal$$publish(promise);
                    }
                    function lib$es6$promise$$internal$$fulfill(promise, value) {
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._result = value;
                        promise._state = lib$es6$promise$$internal$$FULFILLED;
                        if (promise._subscribers.length !== 0) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, promise);
                        }
                    }
                    function lib$es6$promise$$internal$$reject(promise, reason) {
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                            return;
                        }
                        promise._state = lib$es6$promise$$internal$$REJECTED;
                        promise._result = reason;
                        lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publishRejection, promise);
                    }
                    function lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection) {
                        var subscribers = parent._subscribers;
                        var length = subscribers.length;
                        parent._onerror = null;
                        subscribers[length] = child;
                        subscribers[length + lib$es6$promise$$internal$$FULFILLED] = onFulfillment;
                        subscribers[length + lib$es6$promise$$internal$$REJECTED] = onRejection;
                        if (length === 0 && parent._state) {
                            lib$es6$promise$asap$$asap(lib$es6$promise$$internal$$publish, parent);
                        }
                    }
                    function lib$es6$promise$$internal$$publish(promise) {
                        var subscribers = promise._subscribers;
                        var settled = promise._state;
                        if (subscribers.length === 0) {
                            return;
                        }
                        var child, callback, detail = promise._result;
                        for (var i = 0; i < subscribers.length; i += 3) {
                            child = subscribers[i];
                            callback = subscribers[i + settled];
                            if (child) {
                                lib$es6$promise$$internal$$invokeCallback(settled, child, callback, detail);
                            }
                            else {
                                callback(detail);
                            }
                        }
                        promise._subscribers.length = 0;
                    }
                    function lib$es6$promise$$internal$$ErrorObject() {
                        this.error = null;
                    }
                    var lib$es6$promise$$internal$$TRY_CATCH_ERROR = new lib$es6$promise$$internal$$ErrorObject();
                    function lib$es6$promise$$internal$$tryCatch(callback, detail) {
                        try {
                            return callback(detail);
                        }
                        catch (e) {
                            lib$es6$promise$$internal$$TRY_CATCH_ERROR.error = e;
                            return lib$es6$promise$$internal$$TRY_CATCH_ERROR;
                        }
                    }
                    function lib$es6$promise$$internal$$invokeCallback(settled, promise, callback, detail) {
                        var hasCallback = lib$es6$promise$utils$$isFunction(callback), value, error, succeeded, failed;
                        if (hasCallback) {
                            value = lib$es6$promise$$internal$$tryCatch(callback, detail);
                            if (value === lib$es6$promise$$internal$$TRY_CATCH_ERROR) {
                                failed = true;
                                error = value.error;
                                value = null;
                            }
                            else {
                                succeeded = true;
                            }
                            if (promise === value) {
                                lib$es6$promise$$internal$$reject(promise, lib$es6$promise$$internal$$cannotReturnOwn());
                                return;
                            }
                        }
                        else {
                            value = detail;
                            succeeded = true;
                        }
                        if (promise._state !== lib$es6$promise$$internal$$PENDING) {
                        }
                        else if (hasCallback && succeeded) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        else if (failed) {
                            lib$es6$promise$$internal$$reject(promise, error);
                        }
                        else if (settled === lib$es6$promise$$internal$$FULFILLED) {
                            lib$es6$promise$$internal$$fulfill(promise, value);
                        }
                        else if (settled === lib$es6$promise$$internal$$REJECTED) {
                            lib$es6$promise$$internal$$reject(promise, value);
                        }
                    }
                    function lib$es6$promise$$internal$$initializePromise(promise, resolver) {
                        try {
                            resolver(function resolvePromise(value) {
                                lib$es6$promise$$internal$$resolve(promise, value);
                            }, function rejectPromise(reason) {
                                lib$es6$promise$$internal$$reject(promise, reason);
                            });
                        }
                        catch (e) {
                            lib$es6$promise$$internal$$reject(promise, e);
                        }
                    }
                    function lib$es6$promise$enumerator$$Enumerator(Constructor, input) {
                        var enumerator = this;
                        enumerator._instanceConstructor = Constructor;
                        enumerator.promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (enumerator._validateInput(input)) {
                            enumerator._input = input;
                            enumerator.length = input.length;
                            enumerator._remaining = input.length;
                            enumerator._init();
                            if (enumerator.length === 0) {
                                lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                            }
                            else {
                                enumerator.length = enumerator.length || 0;
                                enumerator._enumerate();
                                if (enumerator._remaining === 0) {
                                    lib$es6$promise$$internal$$fulfill(enumerator.promise, enumerator._result);
                                }
                            }
                        }
                        else {
                            lib$es6$promise$$internal$$reject(enumerator.promise, enumerator._validationError());
                        }
                    }
                    lib$es6$promise$enumerator$$Enumerator.prototype._validateInput = function (input) {
                        return lib$es6$promise$utils$$isArray(input);
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._validationError = function () {
                        return new Error('Array Methods must be provided an Array');
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._init = function () {
                        this._result = new Array(this.length);
                    };
                    var lib$es6$promise$enumerator$$default = lib$es6$promise$enumerator$$Enumerator;
                    lib$es6$promise$enumerator$$Enumerator.prototype._enumerate = function () {
                        var enumerator = this;
                        var length = enumerator.length;
                        var promise = enumerator.promise;
                        var input = enumerator._input;
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            enumerator._eachEntry(input[i], i);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._eachEntry = function (entry, i) {
                        var enumerator = this;
                        var c = enumerator._instanceConstructor;
                        if (lib$es6$promise$utils$$isMaybeThenable(entry)) {
                            if (entry.constructor === c && entry._state !== lib$es6$promise$$internal$$PENDING) {
                                entry._onerror = null;
                                enumerator._settledAt(entry._state, i, entry._result);
                            }
                            else {
                                enumerator._willSettleAt(c.resolve(entry), i);
                            }
                        }
                        else {
                            enumerator._remaining--;
                            enumerator._result[i] = entry;
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._settledAt = function (state, i, value) {
                        var enumerator = this;
                        var promise = enumerator.promise;
                        if (promise._state === lib$es6$promise$$internal$$PENDING) {
                            enumerator._remaining--;
                            if (state === lib$es6$promise$$internal$$REJECTED) {
                                lib$es6$promise$$internal$$reject(promise, value);
                            }
                            else {
                                enumerator._result[i] = value;
                            }
                        }
                        if (enumerator._remaining === 0) {
                            lib$es6$promise$$internal$$fulfill(promise, enumerator._result);
                        }
                    };
                    lib$es6$promise$enumerator$$Enumerator.prototype._willSettleAt = function (promise, i) {
                        var enumerator = this;
                        lib$es6$promise$$internal$$subscribe(promise, undefined, function (value) {
                            enumerator._settledAt(lib$es6$promise$$internal$$FULFILLED, i, value);
                        }, function (reason) {
                            enumerator._settledAt(lib$es6$promise$$internal$$REJECTED, i, reason);
                        });
                    };
                    function lib$es6$promise$promise$all$$all(entries) {
                        return new lib$es6$promise$enumerator$$default(this, entries).promise;
                    }
                    var lib$es6$promise$promise$all$$default = lib$es6$promise$promise$all$$all;
                    function lib$es6$promise$promise$race$$race(entries) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        if (!lib$es6$promise$utils$$isArray(entries)) {
                            lib$es6$promise$$internal$$reject(promise, new TypeError('You must pass an array to race.'));
                            return promise;
                        }
                        var length = entries.length;
                        function onFulfillment(value) {
                            lib$es6$promise$$internal$$resolve(promise, value);
                        }
                        function onRejection(reason) {
                            lib$es6$promise$$internal$$reject(promise, reason);
                        }
                        for (var i = 0; promise._state === lib$es6$promise$$internal$$PENDING && i < length; i++) {
                            lib$es6$promise$$internal$$subscribe(Constructor.resolve(entries[i]), undefined, onFulfillment, onRejection);
                        }
                        return promise;
                    }
                    var lib$es6$promise$promise$race$$default = lib$es6$promise$promise$race$$race;
                    function lib$es6$promise$promise$resolve$$resolve(object) {
                        var Constructor = this;
                        if (object && typeof object === 'object' && object.constructor === Constructor) {
                            return object;
                        }
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$resolve(promise, object);
                        return promise;
                    }
                    var lib$es6$promise$promise$resolve$$default = lib$es6$promise$promise$resolve$$resolve;
                    function lib$es6$promise$promise$reject$$reject(reason) {
                        var Constructor = this;
                        var promise = new Constructor(lib$es6$promise$$internal$$noop);
                        lib$es6$promise$$internal$$reject(promise, reason);
                        return promise;
                    }
                    var lib$es6$promise$promise$reject$$default = lib$es6$promise$promise$reject$$reject;
                    var lib$es6$promise$promise$$counter = 0;
                    function lib$es6$promise$promise$$needsResolver() {
                        throw new TypeError('You must pass a resolver function as the first argument to the promise constructor');
                    }
                    function lib$es6$promise$promise$$needsNew() {
                        throw new TypeError("Failed to construct 'Promise': Please use the 'new' operator, this object constructor cannot be called as a function.");
                    }
                    var lib$es6$promise$promise$$default = lib$es6$promise$promise$$Promise;
                    function lib$es6$promise$promise$$Promise(resolver) {
                        this._id = lib$es6$promise$promise$$counter++;
                        this._state = undefined;
                        this._result = undefined;
                        this._subscribers = [];
                        if (lib$es6$promise$$internal$$noop !== resolver) {
                            if (!lib$es6$promise$utils$$isFunction(resolver)) {
                                lib$es6$promise$promise$$needsResolver();
                            }
                            if (!(this instanceof lib$es6$promise$promise$$Promise)) {
                                lib$es6$promise$promise$$needsNew();
                            }
                            lib$es6$promise$$internal$$initializePromise(this, resolver);
                        }
                    }
                    lib$es6$promise$promise$$Promise.all = lib$es6$promise$promise$all$$default;
                    lib$es6$promise$promise$$Promise.race = lib$es6$promise$promise$race$$default;
                    lib$es6$promise$promise$$Promise.resolve = lib$es6$promise$promise$resolve$$default;
                    lib$es6$promise$promise$$Promise.reject = lib$es6$promise$promise$reject$$default;
                    lib$es6$promise$promise$$Promise._setScheduler = lib$es6$promise$asap$$setScheduler;
                    lib$es6$promise$promise$$Promise._setAsap = lib$es6$promise$asap$$setAsap;
                    lib$es6$promise$promise$$Promise._asap = lib$es6$promise$asap$$asap;
                    lib$es6$promise$promise$$Promise.prototype = {
                        constructor: lib$es6$promise$promise$$Promise,
                        then: function (onFulfillment, onRejection) {
                            var parent = this;
                            var state = parent._state;
                            if (state === lib$es6$promise$$internal$$FULFILLED && !onFulfillment || state === lib$es6$promise$$internal$$REJECTED && !onRejection) {
                                return this;
                            }
                            var child = new this.constructor(lib$es6$promise$$internal$$noop);
                            var result = parent._result;
                            if (state) {
                                var callback = arguments[state - 1];
                                lib$es6$promise$asap$$asap(function () {
                                    lib$es6$promise$$internal$$invokeCallback(state, child, callback, result);
                                });
                            }
                            else {
                                lib$es6$promise$$internal$$subscribe(parent, child, onFulfillment, onRejection);
                            }
                            return child;
                        },
                        'catch': function (onRejection) {
                            return this.then(null, onRejection);
                        }
                    };
                    return lib$es6$promise$promise$$default;
                }).call(this);
            }
            PromiseImpl.Init = Init;
        })(PromiseImpl = _Internal.PromiseImpl || (_Internal.PromiseImpl = {}));
    })(_Internal = Office._Internal || (Office._Internal = {}));
    (function (_Internal) {
        function isEdgeLessThan14() {
            var userAgent = window.navigator.userAgent;
            var versionIdx = userAgent.indexOf("Edge/");
            if (versionIdx >= 0) {
                userAgent = userAgent.substring(versionIdx + 5, userAgent.length);
                if (userAgent < "14.14393")
                    return true;
                else
                    return false;
            }
            return false;
        }
        function determinePromise() {
            if (typeof (window) === "undefined" && typeof (Promise) === "function") {
                return Promise;
            }
            if (typeof (window) !== "undefined" && window.Promise) {
                if (isEdgeLessThan14()) {
                    return _Internal.PromiseImpl.Init();
                }
                else {
                    return window.Promise;
                }
            }
            else {
                return _Internal.PromiseImpl.Init();
            }
        }
        _Internal.OfficePromise = determinePromise();
    })(_Internal = Office._Internal || (Office._Internal = {}));
    var OfficePromise = _Internal.OfficePromise;
    Office.Promise = OfficePromise;
})(Office || (Office = {}));
var OTel;
(function (OTel) {
    var CDN_PATH_OTELJS_AGAVE = 'telemetry/oteljs_agave.js';
    var OTelLogger = (function () {
        function OTelLogger() {
        }
        OTelLogger.loaded = function () {
            return !(OTelLogger.logger === undefined || OTelLogger.sink === undefined);
        };
        OTelLogger.getOtelSinkCDNLocation = function () {
            return (OSF._OfficeAppFactory.getLoadScriptHelper().getOfficeJsBasePath() + CDN_PATH_OTELJS_AGAVE);
        };
        OTelLogger.getMapName = function (map, name) {
            if (name !== undefined && map.hasOwnProperty(name)) {
                return map[name];
            }
            return name;
        };
        OTelLogger.getHost = function () {
            var host = OSF._OfficeAppFactory.getHostInfo()["hostType"];
            var map = {
                "excel": "Excel",
                "onenote": "OneNote",
                "outlook": "Outlook",
                "powerpoint": "PowerPoint",
                "project": "Project",
                "visio": "Visio",
                "word": "Word"
            };
            var mappedName = OTelLogger.getMapName(map, host);
            return mappedName;
        };
        OTelLogger.getFlavor = function () {
            var flavor = OSF._OfficeAppFactory.getHostInfo()["hostPlatform"];
            var map = {
                "android": "Android",
                "ios": "iOS",
                "mac": "Mac",
                "universal": "Universal",
                "web": "Web",
                "win32": "Win32"
            };
            var mappedName = OTelLogger.getMapName(map, flavor);
            return mappedName;
        };
        OTelLogger.ensureValue = function (value, alternative) {
            if (!value) {
                return alternative;
            }
            return value;
        };
        OTelLogger.create = function (info) {
            var contract = {
                id: info.appId,
                marketplaceType: info.marketplaceType,
                assetId: info.assetId,
                officeJsVersion: info.officeJSVersion,
                hostJsVersion: info.hostJSVersion,
                browserToken: info.clientId,
                instanceId: info.appInstanceId,
                sessionId: info.sessionId
            };
            var fields = oteljs.Contracts.Office.System.SDX.getFields("SDX", contract);
            var host = OTelLogger.getHost();
            var flavor = OTelLogger.getFlavor();
            var version = (flavor === "Web" && info.hostVersion.slice(0, 2) === "0.") ? "16.0.0.0" : info.hostVersion;
            var context = {
                'App.Name': host,
                'App.Platform': flavor,
                'App.Version': version,
                'Session.Id': OTelLogger.ensureValue(info.correlationId, "00000000-0000-0000-0000-000000000000")
            };
            var namespace = "Office.Extensibility.OfficeJs";
            var ariaTenantToken = 'db334b301e7b474db5e0f02f07c51a47-a1b5bc36-1bbe-482f-a64a-c2d9cb606706-7439';
            var nexusTenantToken = 1755;
            var logger = new oteljs.SimpleTelemetryLogger(undefined, fields);
            logger.setTenantToken(namespace, ariaTenantToken, nexusTenantToken);
            if (oteljs.AgaveSink) {
                OTelLogger.sink = oteljs.AgaveSink.createInstance(context);
            }
            if (OTelLogger.sink === undefined) {
                OTelLogger.attachLegacyAgaveSink(context);
            }
            else {
                logger.addSink(OTelLogger.sink);
            }
            return logger;
        };
        OTelLogger.attachLegacyAgaveSink = function (context) {
            var afterLoadOtelSink = function () {
                if (typeof oteljs_agave !== "undefined") {
                    OTelLogger.sink = oteljs_agave.AgaveSink.createInstance(context);
                }
                if (OTelLogger.sink === undefined || OTelLogger.logger === undefined) {
                    OTelLogger.Enabled = false;
                    OTelLogger.promises = [];
                    OTelLogger.logger = undefined;
                    OTelLogger.sink = undefined;
                    return;
                }
                OTelLogger.logger.addSink(OTelLogger.sink);
                OTelLogger.promises.forEach(function (resolve) {
                    resolve();
                });
            };
            var timeoutAfterFiveSeconds = 5000;
            OSF.OUtil.loadScript(OTelLogger.getOtelSinkCDNLocation(), afterLoadOtelSink, timeoutAfterFiveSeconds);
        };
        OTelLogger.initialize = function (info) {
            if (!OTelLogger.Enabled) {
                OTelLogger.promises = [];
                return;
            }
            var afterOnReady = function () {
                if ((typeof oteljs === "undefined")) {
                    return;
                }
                if (!OTelLogger.loaded()) {
                    OTelLogger.logger = OTelLogger.create(info);
                }
                if (OTelLogger.loaded()) {
                    OTelLogger.promises.forEach(function (resolve) {
                        resolve();
                    });
                }
            };
            Microsoft.Office.WebExtension.onReadyInternal().then(function () { return afterOnReady(); });
        };
        OTelLogger.sendTelemetryEvent = function (telemetryEvent) {
            OTelLogger.onTelemetryLoaded(function () {
                try {
                    OTelLogger.logger.sendTelemetryEvent(telemetryEvent);
                }
                catch (e) {
                }
            });
        };
        OTelLogger.sendCustomerContent = function (customerContentEvent) {
            OTelLogger.onTelemetryLoaded(function () {
                try {
                    OTelLogger.logger.sendCustomerContent(customerContentEvent);
                }
                catch (e) {
                }
            });
        };
        OTelLogger.onTelemetryLoaded = function (resolve) {
            if (!OTelLogger.Enabled) {
                return;
            }
            if (OTelLogger.loaded()) {
                resolve();
            }
            else {
                OTelLogger.promises.push(resolve);
            }
        };
        OTelLogger.promises = [];
        OTelLogger.Enabled = true;
        return OTelLogger;
    }());
    OTel.OTelLogger = OTelLogger;
})(OTel || (OTel = {}));
(function (OfficeExt) {
    var Association = (function () {
        function Association() {
            this.m_mappings = {};
            this.m_onchangeHandlers = [];
        }
        Association.prototype.associate = function (arg1, arg2) {
            function consoleWarn(message) {
                if (typeof console !== 'undefined' && console.warn) {
                    console.warn(message);
                }
            }
            if (arguments.length == 1 && typeof arguments[0] === 'object' && arguments[0]) {
                var mappings = arguments[0];
                for (var key in mappings) {
                    this.associate(key, mappings[key]);
                }
            }
            else if (arguments.length == 2) {
                var name_1 = arguments[0];
                var func = arguments[1];
                if (typeof name_1 !== 'string') {
                    consoleWarn('[InvalidArg] Function=associate');
                    return;
                }
                if (typeof func !== 'function') {
                    consoleWarn('[InvalidArg] Function=associate');
                    return;
                }
                var nameUpperCase = name_1.toUpperCase();
                if (this.m_mappings[nameUpperCase]) {
                    consoleWarn('[DuplicatedName] Function=' + name_1);
                }
                this.m_mappings[nameUpperCase] = func;
                for (var i = 0; i < this.m_onchangeHandlers.length; i++) {
                    this.m_onchangeHandlers[i]();
                }
            }
            else {
                consoleWarn('[InvalidArg] Function=associate');
            }
        };
        Association.prototype.onchange = function (handler) {
            if (handler) {
                this.m_onchangeHandlers.push(handler);
            }
        };
        Object.defineProperty(Association.prototype, "mappings", {
            get: function () {
                return this.m_mappings;
            },
            enumerable: true,
            configurable: true
        });
        return Association;
    }());
    OfficeExt.Association = Association;
})(OfficeExt || (OfficeExt = {}));
var CustomFunctionMappings = window.CustomFunctionMappings || {};
var CustomFunctions;
(function (CustomFunctions) {
    function delayInitialization() {
        CustomFunctionMappings['__delay__'] = true;
    }
    CustomFunctions.delayInitialization = delayInitialization;
    ;
    CustomFunctions._association = new OfficeExt.Association();
    function associate() {
        CustomFunctions._association.associate.apply(CustomFunctions._association, arguments);
        delete CustomFunctionMappings['__delay__'];
    }
    CustomFunctions.associate = associate;
    ;
})(CustomFunctions || (CustomFunctions = {}));
(function (Office) {
    var actions;
    (function (actions) {
        actions._association = new OfficeExt.Association();
        function associate() {
            actions._association.associate.apply(actions._association, arguments);
        }
        actions.associate = associate;
        ;
    })(actions = Office.actions || (Office.actions = {}));
})(Office || (Office = {}));
var g_isExpEnabled = g_isExpEnabled || false;
var g_isOfflineLibrary = g_isOfflineLibrary || false;
(function () {
    var previousConstantNames = OSF.ConstantNames || {};
    OSF.ConstantNames = {
        FileVersion: "16.0.15002.10000",
        OfficeJS: "office.js",
        OfficeDebugJS: "office.debug.js",
        DefaultLocale: "en-us",
        LocaleStringLoadingTimeout: 5000,
        MicrosoftAjaxId: "MSAJAX",
        OfficeStringsId: "OFFICESTRINGS",
        OfficeJsId: "OFFICEJS",
        HostFileId: "HOST",
        O15MappingId: "O15Mapping",
        OfficeStringJS: "office_strings.js",
        O15InitHelper: "o15apptofilemappingtable.js",
        SupportedLocales: OSF.SupportedLocales,
        AssociatedLocales: OSF.AssociatedLocales,
        ExperimentScriptSuffix: "experiment"
    };
    for (var key in previousConstantNames) {
        OSF.ConstantNames[key] = previousConstantNames[key];
    }
})();
OSF.InitializationHelper = function OSF_InitializationHelper(hostInfo, webAppState, context, settings, hostFacade) {
    this._hostInfo = hostInfo;
    this._webAppState = webAppState;
    this._context = context;
    this._settings = settings;
    this._hostFacade = hostFacade;
};
OSF.InitializationHelper.prototype.saveAndSetDialogInfo = function OSF_InitializationHelper$saveAndSetDialogInfo(hostInfoValue) {
};
OSF.InitializationHelper.prototype.getAppContext = function OSF_InitializationHelper$getAppContext(wnd, gotAppContext) {
};
OSF.InitializationHelper.prototype.setAgaveHostCommunication = function OSF_InitializationHelper$setAgaveHostCommunication() {
};
OSF.InitializationHelper.prototype.prepareRightBeforeWebExtensionInitialize = function OSF_InitializationHelper$prepareRightBeforeWebExtensionInitialize(appContext) {
};
OSF.InitializationHelper.prototype.loadAppSpecificScriptAndCreateOM = function OSF_InitializationHelper$loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath) {
};
OSF.InitializationHelper.prototype.prepareRightAfterWebExtensionInitialize = function OSF_InitializationHelper$prepareRightAfterWebExtensionInitialize() {
};
OSF.HostInfoFlags = {
    SharedApp: 1,
    CustomFunction: 2,
    ProtectedDocDisable: 4,
    ExperimentJsEnabled: 8,
    PublicAddin: 0x10
};
OSF._OfficeAppFactory = (function OSF__OfficeAppFactory() {
    var _setNamespace = function OSF_OUtil$_setNamespace(name, parent) {
        if (parent && name && !parent[name]) {
            parent[name] = {};
        }
    };
    _setNamespace("Office", window);
    _setNamespace("Microsoft", window);
    _setNamespace("Office", Microsoft);
    _setNamespace("WebExtension", Microsoft.Office);
    if (typeof (window.Office) === 'object') {
        for (var p in window.Office) {
            if (window.Office.hasOwnProperty(p)) {
                Microsoft.Office.WebExtension[p] = window.Office[p];
            }
        }
    }
    window.Office = Microsoft.Office.WebExtension;
    var initialDisplayModeMappings = {
        0: "Unknown",
        1: "Hidden",
        2: "Taskpane",
        3: "Dialog"
    };
    Microsoft.Office.WebExtension.PlatformType = {
        PC: "PC",
        OfficeOnline: "OfficeOnline",
        Mac: "Mac",
        iOS: "iOS",
        Android: "Android",
        Universal: "Universal"
    };
    Microsoft.Office.WebExtension.HostType = {
        Word: "Word",
        Excel: "Excel",
        PowerPoint: "PowerPoint",
        Outlook: "Outlook",
        OneNote: "OneNote",
        Project: "Project",
        Access: "Access",
        Visio: "Visio"
    };
    var _context = {};
    var _settings = {};
    var _hostFacade = {};
    var _WebAppState = { id: null, webAppUrl: null, conversationID: null, clientEndPoint: null, wnd: window.parent, focused: false };
    var _hostInfo = { isO15: true, isRichClient: true, hostType: "", hostPlatform: "", hostSpecificFileVersion: "", hostLocale: "", osfControlAppCorrelationId: "", isDialog: false, disableLogging: false, flags: 0 };
    var _isLoggingAllowed = true;
    var _initializationHelper = {};
    var _appInstanceId = null;
    var _isOfficeJsLoaded = false;
    var _officeOnReadyPendingResolves = [];
    var _isOfficeOnReadyCalled = false;
    var _officeOnReadyHostAndPlatformInfo = { host: null, platform: null, addin: null };
    var _loadScriptHelper = new ScriptLoading.LoadScriptHelper({
        OfficeJS: OSF.ConstantNames.OfficeJS,
        OfficeDebugJS: OSF.ConstantNames.OfficeDebugJS
    });
    if (window.performance && window.performance.now) {
        _loadScriptHelper.logScriptLoading(OSF.ConstantNames.OfficeJsId, -1, window.performance.now());
    }
    var _windowLocationHash = window.location.hash;
    var _windowLocationSearch = window.location.search;
    var _windowName = window.name;
    var setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks = function (_a) {
        var host = _a.host, platform = _a.platform, addin = _a.addin;
        _isOfficeJsLoaded = true;
        if (typeof OSFPerformance !== "undefined") {
            OSFPerformance.officeOnReady = OSFPerformance.now();
        }
        _officeOnReadyHostAndPlatformInfo = { host: host, platform: platform, addin: addin };
        while (_officeOnReadyPendingResolves.length > 0) {
            _officeOnReadyPendingResolves.shift()(_officeOnReadyHostAndPlatformInfo);
        }
    };
    Microsoft.Office.WebExtension.FeatureGates = {};
    Microsoft.Office.WebExtension.sendTelemetryEvent = function Microsoft_Office_WebExtension_sendTelemetryEvent(telemetryEvent) {
        OTel.OTelLogger.sendTelemetryEvent(telemetryEvent);
    };
    Microsoft.Office.WebExtension.telemetrySink = OTel.OTelLogger;
    Microsoft.Office.WebExtension.onReadyInternal = function Microsoft_Office_WebExtension_onReadyInternal(callback) {
        if (_isOfficeJsLoaded) {
            var host_1 = _officeOnReadyHostAndPlatformInfo.host, platform_1 = _officeOnReadyHostAndPlatformInfo.platform, addin_1 = _officeOnReadyHostAndPlatformInfo.addin;
            if (callback) {
                var result = callback({ host: host_1, platform: platform_1, addin: addin_1 });
                if (result && result.then && typeof result.then === "function") {
                    return result.then(function () { return Office.Promise.resolve({ host: host_1, platform: platform_1, addin: addin_1 }); });
                }
            }
            return Office.Promise.resolve({ host: host_1, platform: platform_1, addin: addin_1 });
        }
        if (callback) {
            return new Office.Promise(function (resolve) {
                _officeOnReadyPendingResolves.push(function (receivedHostAndPlatform) {
                    var result = callback(receivedHostAndPlatform);
                    if (result && result.then && typeof result.then === "function") {
                        return result.then(function () { return resolve(receivedHostAndPlatform); });
                    }
                    resolve(receivedHostAndPlatform);
                });
            });
        }
        return new Office.Promise(function (resolve) {
            _officeOnReadyPendingResolves.push(resolve);
        });
    };
    Microsoft.Office.WebExtension.onReady = function Microsoft_Office_WebExtension_onReady(callback) {
        _isOfficeOnReadyCalled = true;
        return Microsoft.Office.WebExtension.onReadyInternal(callback);
    };
    var getQueryStringValue = function OSF__OfficeAppFactory$getQueryStringValue(paramName) {
        var hostInfoValue;
        var searchString = window.location.search;
        if (searchString) {
            var hostInfoParts = searchString.split(paramName + "=");
            if (hostInfoParts.length > 1) {
                var hostInfoValueRestString = hostInfoParts[1];
                var separatorRegex = new RegExp("[&#]", "g");
                var hostInfoValueParts = hostInfoValueRestString.split(separatorRegex);
                if (hostInfoValueParts.length > 0) {
                    hostInfoValue = hostInfoValueParts[0];
                }
            }
        }
        return hostInfoValue;
    };
    var compareVersions = function _compareVersions(version1, version2) {
        var splitVersion1 = version1.split(".");
        var splitVersion2 = version2.split(".");
        var iter;
        for (iter in splitVersion1) {
            if (parseInt(splitVersion1[iter]) < parseInt(splitVersion2[iter])) {
                return false;
            }
            else if (parseInt(splitVersion1[iter]) > parseInt(splitVersion2[iter])) {
                return true;
            }
        }
        return false;
    };
    var shouldLoadOldOutlookMacJs = function _shouldLoadOldOutlookMacJs() {
        try {
            var versionToUseNewJS = "15.30.1128.0";
            var currentHostVersion = window.external.GetContext().GetHostFullVersion();
        }
        catch (ex) {
            return false;
        }
        return !!compareVersions(versionToUseNewJS, currentHostVersion);
    };
    var _retrieveLoggingAllowed = function OSF__OfficeAppFactory$_retrieveLoggingAllowed() {
        _isLoggingAllowed = true;
        try {
            if (_hostInfo.disableLogging) {
                _isLoggingAllowed = false;
                return;
            }
            window.external = window.external || {};
            if (typeof window.external.GetLoggingAllowed === 'undefined') {
                _isLoggingAllowed = true;
            }
            else {
                _isLoggingAllowed = window.external.GetLoggingAllowed();
            }
        }
        catch (Exception) {
        }
    };
    var _retrieveHostInfo = function OSF__OfficeAppFactory$_retrieveHostInfo() {
        var hostInfoParaName = "_host_Info";
        var hostInfoValue = getQueryStringValue(hostInfoParaName);
        if (!hostInfoValue) {
            try {
                var windowNameObj = JSON.parse(_windowName);
                hostInfoValue = windowNameObj ? windowNameObj["hostInfo"] : null;
            }
            catch (Exception) {
            }
        }
        if (!hostInfoValue) {
            try {
                window.external = window.external || {};
                if (typeof agaveHost !== "undefined" && agaveHost.GetHostInfo) {
                    window.external.GetHostInfo = function () {
                        return agaveHost.GetHostInfo();
                    };
                }
                var fallbackHostInfo = window.external.GetHostInfo();
                if (fallbackHostInfo == "isDialog") {
                    _hostInfo.isO15 = true;
                    _hostInfo.isDialog = true;
                }
                else if (fallbackHostInfo.toLowerCase().indexOf("mac") !== -1 && fallbackHostInfo.toLowerCase().indexOf("outlook") !== -1 && shouldLoadOldOutlookMacJs()) {
                    _hostInfo.isO15 = true;
                }
                else {
                    var hostInfoParts = fallbackHostInfo.split(hostInfoParaName + "=");
                    if (hostInfoParts.length > 1) {
                        hostInfoValue = hostInfoParts[1];
                    }
                    else {
                        hostInfoValue = fallbackHostInfo;
                    }
                }
            }
            catch (Exception) {
            }
        }
        var getSessionStorage = function OSF__OfficeAppFactory$_retrieveHostInfo$getSessionStorage() {
            var osfSessionStorage = null;
            try {
                if (window.sessionStorage) {
                    osfSessionStorage = window.sessionStorage;
                }
            }
            catch (ex) {
            }
            return osfSessionStorage;
        };
        var osfSessionStorage = getSessionStorage();
        if (!hostInfoValue && osfSessionStorage && osfSessionStorage.getItem("hostInfoValue")) {
            hostInfoValue = osfSessionStorage.getItem("hostInfoValue");
        }
        if (hostInfoValue) {
            hostInfoValue = decodeURIComponent(hostInfoValue);
            _hostInfo.isO15 = false;
            var items = hostInfoValue.split("$");
            if (typeof items[2] == "undefined") {
                items = hostInfoValue.split("|");
            }
            _hostInfo.hostType = (typeof items[0] == "undefined") ? "" : items[0].toLowerCase();
            _hostInfo.hostPlatform = (typeof items[1] == "undefined") ? "" : items[1].toLowerCase();
            ;
            _hostInfo.hostSpecificFileVersion = (typeof items[2] == "undefined") ? "" : items[2].toLowerCase();
            _hostInfo.hostLocale = (typeof items[3] == "undefined") ? "" : items[3].toLowerCase();
            _hostInfo.osfControlAppCorrelationId = (typeof items[4] == "undefined") ? "" : items[4];
            if (_hostInfo.osfControlAppCorrelationId == "telemetry") {
                _hostInfo.osfControlAppCorrelationId = "";
            }
            _hostInfo.isDialog = (((typeof items[5]) != "undefined") && items[5] == "isDialog") ? true : false;
            _hostInfo.disableLogging = (((typeof items[6]) != "undefined") && items[6] == "disableLogging") ? true : false;
            _hostInfo.flags = (((typeof items[7]) === "string") && items[7].length > 0) ? parseInt(items[7]) : 0;
            if (g_isOfflineLibrary) {
                g_isExpEnabled = false;
            }
            else {
                g_isExpEnabled = g_isExpEnabled || !!(_hostInfo.flags & OSF.HostInfoFlags.ExperimentJsEnabled);
            }
            var hostSpecificFileVersionValue = parseFloat(_hostInfo.hostSpecificFileVersion);
            var fallbackVersion = OSF.HostSpecificFileVersionDefault;
            if (OSF.HostSpecificFileVersionMap[_hostInfo.hostType] && OSF.HostSpecificFileVersionMap[_hostInfo.hostType][_hostInfo.hostPlatform]) {
                fallbackVersion = OSF.HostSpecificFileVersionMap[_hostInfo.hostType][_hostInfo.hostPlatform];
            }
            if (hostSpecificFileVersionValue > parseFloat(fallbackVersion)) {
                _hostInfo.hostSpecificFileVersion = fallbackVersion;
            }
            if (osfSessionStorage) {
                try {
                    osfSessionStorage.setItem("hostInfoValue", hostInfoValue);
                }
                catch (e) {
                }
            }
        }
        else {
            _hostInfo.isO15 = true;
            _hostInfo.hostLocale = getQueryStringValue("locale");
        }
    };
    var getAppContextAsync = function OSF__OfficeAppFactory$getAppContextAsync(wnd, gotAppContext) {
        if (OSF.AppTelemetry && OSF.AppTelemetry.logAppCommonMessage) {
            OSF.AppTelemetry.logAppCommonMessage("getAppContextAsync starts");
        }
        _initializationHelper.getAppContext(wnd, gotAppContext);
    };
    var initialize = function OSF__OfficeAppFactory$initialize() {
        _retrieveHostInfo();
        _retrieveLoggingAllowed();
        if (_hostInfo.hostPlatform == "web" && _hostInfo.isDialog && window == window.top && window.opener == null) {
            window.open('', '_self', '');
            window.close();
        }
        if ((_hostInfo.flags & (OSF.HostInfoFlags.SharedApp | OSF.HostInfoFlags.CustomFunction)) !== 0) {
            if (typeof (window.Promise) === 'undefined') {
                window.Promise = window.Office.Promise;
            }
        }
        _loadScriptHelper.setAppCorrelationId(_hostInfo.osfControlAppCorrelationId);
        var basePath = _loadScriptHelper.getOfficeJsBasePath();
        var requiresMsAjax = false;
        if (!basePath)
            throw "Office Web Extension script library file name should be " + OSF.ConstantNames.OfficeJS + " or " + OSF.ConstantNames.OfficeDebugJS + ".";
        var isMicrosftAjaxLoaded = function OSF$isMicrosftAjaxLoaded() {
            if ((typeof (Sys) !== 'undefined' && typeof (Type) !== 'undefined' &&
                Sys.StringBuilder && typeof (Sys.StringBuilder) === "function" &&
                Type.registerNamespace && typeof (Type.registerNamespace) === "function" &&
                Type.registerClass && typeof (Type.registerClass) === "function") ||
                (typeof (OfficeExt) !== "undefined" && OfficeExt.MsAjaxError)) {
                return true;
            }
            else {
                return false;
            }
        };
        var officeStrings = null;
        var loadLocaleStrings = function OSF__OfficeAppFactory_initialize$loadLocaleStrings(appLocale) {
            var fallbackLocaleTried = false;
            var loadLocaleStringCallback = function OSF__OfficeAppFactory_initialize$loadLocaleStringCallback() {
                if (typeof Strings == 'undefined' || typeof Strings.OfficeOM == 'undefined') {
                    if (!fallbackLocaleTried) {
                        fallbackLocaleTried = true;
                        var fallbackLocaleStringFile = basePath + OSF.ConstantNames.DefaultLocale + "/" + OSF.ConstantNames.OfficeStringJS;
                        _loadScriptHelper.loadScript(fallbackLocaleStringFile, OSF.ConstantNames.OfficeStringsId, loadLocaleStringCallback, true, OSF.ConstantNames.LocaleStringLoadingTimeout);
                        return false;
                    }
                    else {
                        throw "Neither the locale, " + appLocale.toLowerCase() + ", provided by the host app nor the fallback locale " + OSF.ConstantNames.DefaultLocale + " are supported.";
                    }
                }
                else {
                    fallbackLocaleTried = false;
                    officeStrings = Strings.OfficeOM;
                }
            };
            if (!isMicrosftAjaxLoaded()) {
                window.Type = Function;
                Type.registerNamespace = function (ns) {
                    window[ns] = window[ns] || {};
                };
                Type.prototype.registerClass = function (cls) {
                    cls = {};
                };
            }
            var localeStringFile = basePath + OSF.getSupportedLocale(appLocale, OSF.ConstantNames.DefaultLocale) + "/" + OSF.ConstantNames.OfficeStringJS;
            _loadScriptHelper.loadScript(localeStringFile, OSF.ConstantNames.OfficeStringsId, loadLocaleStringCallback, true, OSF.ConstantNames.LocaleStringLoadingTimeout);
        };
        var onAppCodeAndMSAjaxReady = function OSF__OfficeAppFactory_initialize$onAppCodeAndMSAjaxReady(loadSuccess) {
            if (loadSuccess) {
                _initializationHelper = new OSF.InitializationHelper(_hostInfo, _WebAppState, _context, _settings, _hostFacade);
                if (_hostInfo.hostPlatform == "web" && _initializationHelper.saveAndSetDialogInfo) {
                    _initializationHelper.saveAndSetDialogInfo(getQueryStringValue("_host_Info"));
                }
                _initializationHelper.setAgaveHostCommunication();
                if (typeof OSFPerformance !== "undefined") {
                    OSFPerformance.getAppContextStart = OSFPerformance.now();
                }
                getAppContextAsync(_WebAppState.wnd, function (appContext) {
                    if (typeof OSFPerformance !== "undefined") {
                        OSFPerformance.getAppContextEnd = OSFPerformance.now();
                    }
                    if (OSF.AppTelemetry && OSF.AppTelemetry.logAppCommonMessage) {
                        OSF.AppTelemetry.logAppCommonMessage("getAppContextAsync callback start");
                    }
                    _appInstanceId = appContext._appInstanceId;
                    if (appContext.get_featureGates) {
                        var featureGates = appContext.get_featureGates();
                        if (featureGates) {
                            Microsoft.Office.WebExtension.FeatureGates = featureGates;
                        }
                    }
                    var updateVersionInfo = function updateVersionInfo() {
                        var hostVersionItems = _hostInfo.hostSpecificFileVersion.split(".");
                        if (appContext.get_appMinorVersion) {
                            var isIOS = _hostInfo.hostPlatform == "ios";
                            if (!isIOS) {
                                if (isNaN(appContext.get_appMinorVersion())) {
                                    appContext._appMinorVersion = parseInt(hostVersionItems[1]);
                                }
                                else if (hostVersionItems.length > 1 && !isNaN(Number(hostVersionItems[1]))) {
                                    appContext._appMinorVersion = parseInt(hostVersionItems[1]);
                                }
                            }
                        }
                        if (_hostInfo.isDialog) {
                            appContext._isDialog = _hostInfo.isDialog;
                        }
                    };
                    updateVersionInfo();
                    var appReady = function appReady() {
                        _initializationHelper.prepareApiSurface && _initializationHelper.prepareApiSurface(appContext);
                        _loadScriptHelper.waitForFunction(function () { return (Microsoft.Office.WebExtension.initialize != undefined || _isOfficeOnReadyCalled); }, function (initializedDeclaredOrOfficeOnReadyCalled) {
                            if (initializedDeclaredOrOfficeOnReadyCalled) {
                                if (_initializationHelper.prepareApiSurface) {
                                    if (Microsoft.Office.WebExtension.initialize) {
                                        Microsoft.Office.WebExtension.initialize(_initializationHelper.getInitializationReason(appContext));
                                    }
                                }
                                else {
                                    if (!Microsoft.Office.WebExtension.initialize) {
                                        Microsoft.Office.WebExtension.initialize = function () { };
                                    }
                                    _initializationHelper.prepareRightBeforeWebExtensionInitialize(appContext);
                                }
                                _initializationHelper.prepareRightAfterWebExtensionInitialize && _initializationHelper.prepareRightAfterWebExtensionInitialize(appContext);
                                var appNumber = appContext.get_appName();
                                var addinInfo = null;
                                if ((_hostInfo.flags & OSF.HostInfoFlags.SharedApp) !== 0) {
                                    addinInfo = {
                                        visibilityMode: initialDisplayModeMappings[(appContext.get_initialDisplayMode && typeof appContext.get_initialDisplayMode === 'function') ? appContext.get_initialDisplayMode() : 0]
                                    };
                                }
                                setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks({
                                    host: OfficeExt.HostName.Host.getInstance().getHost(appNumber),
                                    platform: OfficeExt.HostName.Host.getInstance().getPlatform(appNumber),
                                    addin: addinInfo
                                });
                            }
                            else {
                                throw new Error("Office.js has not fully loaded. Your app must call \"Office.onReady()\" as part of it's loading sequence (or set the \"Office.initialize\" function). If your app has this functionality, try reloading this page.");
                            }
                        }, 400, 50);
                    };
                    if (!_loadScriptHelper.isScriptLoading(OSF.ConstantNames.OfficeStringsId)) {
                        loadLocaleStrings(appContext.get_appUILocale());
                    }
                    _loadScriptHelper.waitForScripts([OSF.ConstantNames.OfficeStringsId], function () {
                        if (officeStrings && !Strings.OfficeOM) {
                            Strings.OfficeOM = officeStrings;
                        }
                        _initializationHelper.loadAppSpecificScriptAndCreateOM(appContext, appReady, basePath);
                        if (typeof OSFPerformance !== "undefined") {
                            OSFPerformance.createOMEnd = OSFPerformance.now();
                        }
                    });
                });
                if (_hostInfo.isO15) {
                    var wacXdmInfoIsMissing = (OSF.OUtil.parseXdmInfo() == null);
                    if (wacXdmInfoIsMissing) {
                        var isPlainBrowser = true;
                        if (window.external && typeof window.external.GetContext !== 'undefined') {
                            try {
                                window.external.GetContext();
                                isPlainBrowser = false;
                            }
                            catch (e) {
                            }
                        }
                        if (typeof OsfOptOut === "undefined" && isPlainBrowser && window.top !== window.self) {
                            if (window.console && window.console.log) {
                                window.console.log("The add-in is not hosted in plain browser top window.");
                            }
                            window.location.href = "about:blank";
                        }
                        if (isPlainBrowser) {
                            setOfficeJsAsLoadedAndDispatchPendingOnReadyCallbacks({
                                host: null,
                                platform: null,
                                addin: null
                            });
                        }
                    }
                }
            }
            else {
                var errorMsg = "MicrosoftAjax.js is not loaded successfully.";
                if (OSF.AppTelemetry && OSF.AppTelemetry.logAppException) {
                    OSF.AppTelemetry.logAppException(errorMsg);
                }
                throw errorMsg;
            }
        };
        var onAppCodeReady = function OSF__OfficeAppFactory_initialize$onAppCodeReady() {
            if (OSF.AppTelemetry && OSF.AppTelemetry.setOsfControlAppCorrelationId) {
                OSF.AppTelemetry.setOsfControlAppCorrelationId(_hostInfo.osfControlAppCorrelationId);
            }
            if (_loadScriptHelper.isScriptLoading(OSF.ConstantNames.MicrosoftAjaxId)) {
                _loadScriptHelper.waitForScripts([OSF.ConstantNames.MicrosoftAjaxId], onAppCodeAndMSAjaxReady);
            }
            else {
                _loadScriptHelper.waitForFunction(isMicrosftAjaxLoaded, onAppCodeAndMSAjaxReady, 500, 100);
            }
        };
        if (_hostInfo.isO15) {
            _loadScriptHelper.loadScript(basePath + OSF.ConstantNames.O15InitHelper, OSF.ConstantNames.O15MappingId, onAppCodeReady);
        }
        else {
            var hostSpecificFileName;
            hostSpecificFileName = ([
                _hostInfo.hostType,
                _hostInfo.hostPlatform,
                _hostInfo.hostSpecificFileVersion,
                OSF.ConstantNames.HostFileScriptSuffix || null,
            ]
                .filter(function (part) { return part != null; })
                .join("-"))
                +
                    ".js";
            _loadScriptHelper.loadScript(basePath + hostSpecificFileName.toLowerCase(), OSF.ConstantNames.HostFileId, onAppCodeReady);
            if (typeof OSFPerformance !== "undefined") {
                OSFPerformance.hostSpecificFileName = hostSpecificFileName.toLowerCase();
            }
        }
        if (_hostInfo.hostLocale) {
            loadLocaleStrings(_hostInfo.hostLocale);
        }
        if (requiresMsAjax && !isMicrosftAjaxLoaded()) {
            var msAjaxCDNPath = (window.location.protocol.toLowerCase() === 'https:' ? 'https:' : 'http:') + '//ajax.aspnetcdn.com/ajax/3.5/MicrosoftAjax.js';
            _loadScriptHelper.loadScriptParallel(msAjaxCDNPath, OSF.ConstantNames.MicrosoftAjaxId);
        }
        window.confirm = function OSF__OfficeAppFactory_initialize$confirm(message) {
            throw new Error('Function window.confirm is not supported.');
        };
        window.alert = function OSF__OfficeAppFactory_initialize$alert(message) {
            throw new Error('Function window.alert is not supported.');
        };
        window.prompt = function OSF__OfficeAppFactory_initialize$prompt(message, defaultvalue) {
            throw new Error('Function window.prompt is not supported.');
        };
        var isOutlookAndroid = _hostInfo.hostType == "outlook" && _hostInfo.hostPlatform == "android";
        if (!isOutlookAndroid) {
            window.history.replaceState = null;
            window.history.pushState = null;
        }
    };
    initialize();
    if (window.addEventListener) {
        window.addEventListener('DOMContentLoaded', function (event) {
            Microsoft.Office.WebExtension.onReadyInternal(function () {
                if (typeof OSFPerfUtil !== 'undefined') {
                    OSFPerfUtil.sendPerformanceTelemetry();
                }
            });
        });
    }
    return {
        getId: function OSF__OfficeAppFactory$getId() { return _WebAppState.id; },
        getClientEndPoint: function OSF__OfficeAppFactory$getClientEndPoint() { return _WebAppState.clientEndPoint; },
        getContext: function OSF__OfficeAppFactory$getContext() { return _context; },
        setContext: function OSF__OfficeAppFactory$setContext(context) { _context = context; },
        getHostInfo: function OSF_OfficeAppFactory$getHostInfo() { return _hostInfo; },
        getLoggingAllowed: function OSF_OfficeAppFactory$getLoggingAllowed() { return _isLoggingAllowed; },
        getHostFacade: function OSF__OfficeAppFactory$getHostFacade() { return _hostFacade; },
        setHostFacade: function setHostFacade(hostFacade) { _hostFacade = hostFacade; },
        getInitializationHelper: function OSF__OfficeAppFactory$getInitializationHelper() { return _initializationHelper; },
        getCachedSessionSettingsKey: function OSF__OfficeAppFactory$getCachedSessionSettingsKey() {
            return (_WebAppState.conversationID != null ? _WebAppState.conversationID : _appInstanceId) + "CachedSessionSettings";
        },
        getWebAppState: function OSF__OfficeAppFactory$getWebAppState() { return _WebAppState; },
        getWindowLocationHash: function OSF__OfficeAppFactory$getHash() { return _windowLocationHash; },
        getWindowLocationSearch: function OSF__OfficeAppFactory$getSearch() { return _windowLocationSearch; },
        getLoadScriptHelper: function OSF__OfficeAppFactory$getLoadScriptHelper() { return _loadScriptHelper; },
        getWindowName: function OSF__OfficeAppFactory$getWindowName() { return _windowName; }
    };
})();



var oteljs=function(t){var e={};function n(r){if(e[r])return e[r].exports;var i=e[r]={i:r,l:!1,exports:{}};return t[r].call(i.exports,i,i.exports,n),i.l=!0,i.exports}return n.m=t,n.c=e,n.d=function(t,e,r){n.o(t,e)||Object.defineProperty(t,e,{enumerable:!0,get:r})},n.r=function(t){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(t,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(t,"__esModule",{value:!0})},n.t=function(t,e){if(1&e&&(t=n(t)),8&e)return t;if(4&e&&"object"==typeof t&&t&&t.__esModule)return t;var r=Object.create(null);if(n.r(r),Object.defineProperty(r,"default",{enumerable:!0,value:t}),2&e&&"string"!=typeof t)for(var i in t)n.d(r,i,function(e){return t[e]}.bind(null,i));return r},n.n=function(t){var e=t&&t.__esModule?function(){return t.default}:function(){return t};return n.d(e,"a",e),e},n.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},n.p="",n(n.s=19)}([function(t,e,n){"use strict";n.d(e,"a",(function(){return o})),n.d(e,"d",(function(){return a})),n.d(e,"b",(function(){return c})),n.d(e,"e",(function(){return s})),n.d(e,"c",(function(){return u}));var r=n(3),i=n(4);function o(t,e){return{name:t,dataType:r.a.Boolean,value:e,classification:i.a.SystemMetadata}}function a(t,e){return{name:t,dataType:r.a.Int64,value:e,classification:i.a.SystemMetadata}}function c(t,e){return{name:t,dataType:r.a.Double,value:e,classification:i.a.SystemMetadata}}function s(t,e){return{name:t,dataType:r.a.String,value:e,classification:i.a.SystemMetadata}}function u(t,e){return{name:t,dataType:r.a.Guid,value:e,classification:i.a.SystemMetadata}}},function(t,e,n){"use strict";n.d(e,"b",(function(){return r})),n.d(e,"a",(function(){return i})),n.d(e,"e",(function(){return a})),n.d(e,"d",(function(){return c})),n.d(e,"c",(function(){return s}));var r,i,o=new(n(10).a);function a(){return o}function c(t,e,n){o.fireEvent({level:t,category:e,message:n})}function s(t,e,n){c(r.Error,t,(function(){var t=n instanceof Error?n.message:"";return e+": "+t}))}!function(t){t[t.Error=0]="Error",t[t.Warning=1]="Warning",t[t.Info=2]="Info",t[t.Verbose=3]="Verbose"}(r||(r={})),function(t){t[t.Core=0]="Core",t[t.Sink=1]="Sink",t[t.Transport=2]="Transport"}(i||(i={}))},function(t,e,n){"use strict";n.d(e,"a",(function(){return i}));var r=n(0);function i(t,e,n){t.push(Object(r.e)("zC."+e,n))}},function(t,e,n){"use strict";var r;n.d(e,"a",(function(){return r})),function(t){t[t.String=0]="String",t[t.Boolean=1]="Boolean",t[t.Int64=2]="Int64",t[t.Double=3]="Double",t[t.Guid=4]="Guid"}(r||(r={}))},function(t,e,n){"use strict";var r;n.d(e,"a",(function(){return r})),function(t){t[t.EssentialServiceMetadata=1]="EssentialServiceMetadata",t[t.AccountData=2]="AccountData",t[t.SystemMetadata=4]="SystemMetadata",t[t.OrganizationIdentifiableInformation=8]="OrganizationIdentifiableInformation",t[t.EndUserIdentifiableInformation=16]="EndUserIdentifiableInformation",t[t.CustomerContent=32]="CustomerContent",t[t.AccessControl=64]="AccessControl"}(r||(r={}))},function(t,e,n){"use strict";var r,i,o,a,c;n.d(e,"e",(function(){return r})),n.d(e,"d",(function(){return i})),n.d(e,"a",(function(){return o})),n.d(e,"b",(function(){return a})),n.d(e,"c",(function(){return c})),function(t){t[t.NotSet=0]="NotSet",t[t.Measure=1]="Measure",t[t.Diagnostics=2]="Diagnostics",t[t.CriticalBusinessImpact=191]="CriticalBusinessImpact",t[t.CriticalCensus=192]="CriticalCensus",t[t.CriticalExperimentation=193]="CriticalExperimentation",t[t.CriticalUsage=194]="CriticalUsage"}(r||(r={})),function(t){t[t.NotSet=0]="NotSet",t[t.Normal=1]="Normal",t[t.High=2]="High"}(i||(i={})),function(t){t[t.NotSet=0]="NotSet",t[t.Normal=1]="Normal",t[t.High=2]="High"}(o||(o={})),function(t){t[t.NotSet=0]="NotSet",t[t.SoftwareSetup=1]="SoftwareSetup",t[t.ProductServiceUsage=2]="ProductServiceUsage",t[t.ProductServicePerformance=4]="ProductServicePerformance",t[t.DeviceConfiguration=8]="DeviceConfiguration",t[t.InkingTypingSpeech=16]="InkingTypingSpeech"}(a||(a={})),function(t){t[t.ReservedDoNotUse=0]="ReservedDoNotUse",t[t.BasicEvent=10]="BasicEvent",t[t.FullEvent=100]="FullEvent",t[t.NecessaryServiceDataEvent=110]="NecessaryServiceDataEvent",t[t.AlwaysOnNecessaryServiceDataEvent=120]="AlwaysOnNecessaryServiceDataEvent"}(c||(c={}))},function(t,e,n){"use strict";n.d(e,"a",(function(){return p}));var r,i,o,a,c,s,u,d,f,l=n(0),v=n(2);(r||(r={})).getFields=function(t,e){var n=[];return n.push(Object(l.d)(t+".Code",e.code)),void 0!==e.type&&n.push(Object(l.e)(t+".Type",e.type)),void 0!==e.tag&&n.push(Object(l.d)(t+".Tag",e.tag)),void 0!==e.isExpected&&n.push(Object(l.a)(t+".IsExpected",e.isExpected)),Object(v.a)(n,t,"Office.System.Result"),n},(o=i||(i={})).contractName="Office.System.Activity",o.getFields=function(t){var e=[];return void 0!==t.cV&&e.push(Object(l.e)("Activity.CV",t.cV)),e.push(Object(l.d)("Activity.Duration",t.duration)),e.push(Object(l.d)("Activity.Count",t.count)),e.push(Object(l.d)("Activity.AggMode",t.aggMode)),void 0!==t.success&&e.push(Object(l.a)("Activity.Success",t.success)),void 0!==t.result&&e.push.apply(e,r.getFields("Activity.Result",t.result)),Object(v.a)(e,"Activity",o.contractName),e},(a||(a={})).getFields=function(t,e){var n=[];return void 0!==e.id&&n.push(Object(l.e)(t+".Id",e.id)),void 0!==e.version&&n.push(Object(l.e)(t+".Version",e.version)),void 0!==e.sessionId&&n.push(Object(l.e)(t+".SessionId",e.sessionId)),Object(v.a)(n,t,"Office.System.Host"),n},(c||(c={})).getFields=function(t,e){var n=[];return void 0!==e.alias&&n.push(Object(l.e)(t+".Alias",e.alias)),void 0!==e.primaryIdentityHash&&n.push(Object(l.e)(t+".PrimaryIdentityHash",e.primaryIdentityHash)),void 0!==e.primaryIdentitySpace&&n.push(Object(l.e)(t+".PrimaryIdentitySpace",e.primaryIdentitySpace)),void 0!==e.tenantId&&n.push(Object(l.e)(t+".TenantId",e.tenantId)),void 0!==e.tenantGroup&&n.push(Object(l.e)(t+".TenantGroup",e.tenantGroup)),void 0!==e.isAnonymous&&n.push(Object(l.a)(t+".IsAnonymous",e.isAnonymous)),Object(v.a)(n,t,"Office.System.User"),n},(s||(s={})).getFields=function(t,e){var n=[];return void 0!==e.id&&n.push(Object(l.e)(t+".Id",e.id)),void 0!==e.version&&n.push(Object(l.e)(t+".Version",e.version)),void 0!==e.instanceId&&n.push(Object(l.e)(t+".InstanceId",e.instanceId)),void 0!==e.name&&n.push(Object(l.e)(t+".Name",e.name)),void 0!==e.marketplaceType&&n.push(Object(l.e)(t+".MarketplaceType",e.marketplaceType)),void 0!==e.sessionId&&n.push(Object(l.e)(t+".SessionId",e.sessionId)),void 0!==e.browserToken&&n.push(Object(l.e)(t+".BrowserToken",e.browserToken)),void 0!==e.osfRuntimeVersion&&n.push(Object(l.e)(t+".OsfRuntimeVersion",e.osfRuntimeVersion)),void 0!==e.officeJsVersion&&n.push(Object(l.e)(t+".OfficeJsVersion",e.officeJsVersion)),void 0!==e.hostJsVersion&&n.push(Object(l.e)(t+".HostJsVersion",e.hostJsVersion)),void 0!==e.assetId&&n.push(Object(l.e)(t+".AssetId",e.assetId)),void 0!==e.providerName&&n.push(Object(l.e)(t+".ProviderName",e.providerName)),void 0!==e.type&&n.push(Object(l.e)(t+".Type",e.type)),Object(v.a)(n,t,"Office.System.SDX"),n},(u||(u={})).getFields=function(t,e){var n=[];return void 0!==e.name&&n.push(Object(l.e)(t+".Name",e.name)),void 0!==e.state&&n.push(Object(l.e)(t+".State",e.state)),Object(v.a)(n,t,"Office.System.Funnel"),n},(d||(d={})).getFields=function(t,e){var n=[];return void 0!==e.id&&n.push(Object(l.d)(t+".Id",e.id)),void 0!==e.name&&n.push(Object(l.e)(t+".Name",e.name)),void 0!==e.commandSurface&&n.push(Object(l.e)(t+".CommandSurface",e.commandSurface)),void 0!==e.parentName&&n.push(Object(l.e)(t+".ParentName",e.parentName)),void 0!==e.triggerMethod&&n.push(Object(l.e)(t+".TriggerMethod",e.triggerMethod)),void 0!==e.timeOffsetMs&&n.push(Object(l.d)(t+".TimeOffsetMs",e.timeOffsetMs)),Object(v.a)(n,t,"Office.System.UserAction"),n},function(t){t.getFields=function(t,e){var n=[];return n.push(Object(l.e)(t+".ErrorGroup",e.errorGroup)),n.push(Object(l.d)(t+".Tag",e.tag)),void 0!==e.code&&n.push(Object(l.d)(t+".Code",e.code)),void 0!==e.id&&n.push(Object(l.d)(t+".Id",e.id)),void 0!==e.count&&n.push(Object(l.d)(t+".Count",e.count)),Object(v.a)(n,t,"Office.System.Error"),n}}(f||(f={}));var p,y=i,h=r,g=f,m=u,b=a,F=s,O=d,S=c;!function(t){!function(t){!function(t){t.Activity=y,t.Result=h,t.Error=g,t.Funnel=m,t.Host=b,t.SDX=F,t.User=S,t.UserAction=O}(t.System||(t.System={}))}(t.Office||(t.Office={}))}(p||(p={}))},function(t,e,n){"use strict";function r(t){var e={eventName:t.eventName,eventFlags:t.eventFlags};return t.telemetryProperties&&(e.telemetryProperties={ariaTenantToken:t.telemetryProperties.ariaTenantToken,nexusTenantToken:t.telemetryProperties.nexusTenantToken}),t.eventContract&&(e.eventContract={name:t.eventContract.name,dataFields:t.eventContract.dataFields.slice()}),e.dataFields=t.dataFields?t.dataFields.slice():[],e}n.d(e,"a",(function(){return r}))},function(t,e,n){"use strict";n.d(e,"b",(function(){return l})),n.d(e,"a",(function(){return v}));var r,i,o=n(7),a=n(1);!function(t){t[t.Aria=0]="Aria",t[t.Nexus=1]="Nexus"}(r||(r={})),function(t){var e={},n={},i={};function o(t){if("object"!=typeof t)throw new Error("tokenTree must be an object");i=function t(e,n){if("object"!=typeof n)return n;for(var r=0,i=Object.keys(n);r<i.length;r++){var o=i[r];o in e&&(e[o],1)?e[o]=t(e[o],n[o]):e[o]=n[o]}return e}(i,t)}function c(t){if(e[t])return e[t];var n=u(t,r.Aria);return"string"==typeof n?(e[t]=n,n):void 0}function s(t){if(n[t])return n[t];var e=u(t,r.Nexus);return"number"==typeof e?(n[t]=e,e):void 0}function u(t,e){var n=t.split("."),o=i,a=void 0;if(o){for(var c=0;c<n.length-1;c++)o[n[c]]&&(o=o[n[c]],e===r.Aria&&"string"==typeof o.ariaTenantToken?a=o.ariaTenantToken:e===r.Nexus&&"number"==typeof o.nexusTenantToken&&(a=o.nexusTenantToken));return a}}t.setTenantToken=function(t,e,n){var r=t.split(".");if(r.length<2||"Office"!==r[0])Object(a.d)(a.b.Error,a.a.Core,(function(){return"Invalid namespace: "+t}));else{var i=Object.create(Object.prototype);e&&(i.ariaTenantToken=e),n&&(i.nexusTenantToken=n);var c,s=i;for(c=r.length-1;c>=0;--c){var u=Object.create(Object.prototype);u[r[c]]=s,s=u}o(s)}},t.setTenantTokens=o,t.getTenantTokens=function(t){var e=c(t),n=s(t);if(!n||!e)throw new Error("Could not find tenant token for "+t);return{ariaTenantToken:e,nexusTenantToken:n}},t.getAriaTenantToken=c,t.getNexusTenantToken=s,t.clear=function(){e={},n={},i={}}}(i||(i={}));var c,s=n(3);!function(t){var e=/^[A-Z][a-zA-Z0-9]*$/,n=/^[a-zA-Z0-9_\.]*$/;function r(t){return void 0!==t&&n.test(t)}function i(t){if(!((e=t.name)&&r(e)&&e.length+5<100))throw new Error("Invalid dataField name");var e;t.dataType===s.a.Int64&&o(t.value)}function o(t){if("number"!=typeof t||!isFinite(t)||Math.floor(t)!==t||t<-9007199254740991||t>9007199254740991)throw new Error("Invalid integer "+JSON.stringify(t))}t.validateTelemetryEvent=function(t){if(!function(t){if(!t||t.length>98)return!1;var n=t.split("."),r=n[n.length-1];return function(t){return!!t&&t.length>=3&&"Office"===t[0]}(n)&&(i=r,void 0!==i&&e.test(i));var i}(t.eventName))throw new Error("Invalid eventName");if(t.eventContract&&!r(t.eventContract.name))throw new Error("Invalid eventContract");if(null!=t.dataFields)for(var n=0;n<t.dataFields.length;n++)i(t.dataFields[n])},t.validateInt=o}(c||(c={}));var u=n(10),d=n(0),f=function(){return(f=Object.assign||function(t){for(var e,n=1,r=arguments.length;n<r;n++)for(var i in e=arguments[n])Object.prototype.hasOwnProperty.call(e,i)&&(t[i]=e[i]);return t}).apply(this,arguments)},l=-1,v=function(){function t(t,e,n){var r,i;this.onSendEvent=new u.a,this.persistentDataFields=[],this.config=n||{},t&&(this.onSendEvent=t.onSendEvent,(r=this.persistentDataFields).push.apply(r,t.persistentDataFields),this.config=f(f({},t.getConfig()),this.config)),e&&(i=this.persistentDataFields).push.apply(i,e)}return t.prototype.sendTelemetryEvent=function(t){var e;try{if(0===this.onSendEvent.getListenerCount())return void Object(a.d)(a.b.Warning,a.a.Core,(function(){return"No telemetry sinks are attached."}));e=this.cloneEvent(t),this.processTelemetryEvent(e)}catch(t){return void Object(a.c)(a.a.Core,"SendTelemetryEvent",t)}try{this.onSendEvent.fireEvent(e)}catch(t){}},t.prototype.processTelemetryEvent=function(t){var e;t.telemetryProperties||(t.telemetryProperties=i.getTenantTokens(t.eventName)),t.dataFields&&(t.dataFields.unshift(Object(d.e)("OTelJS.Version","3.1.74")),this.persistentDataFields&&(e=t.dataFields).unshift.apply(e,this.persistentDataFields)),this.config.disableValidation||c.validateTelemetryEvent(t)},t.prototype.addSink=function(t){this.onSendEvent.addListener((function(e){return t.sendTelemetryEvent(e)}))},t.prototype.setTenantToken=function(t,e,n){i.setTenantToken(t,e,n)},t.prototype.setTenantTokens=function(t){i.setTenantTokens(t)},t.prototype.cloneEvent=function(t){return Object(o.a)(t)},t.prototype.getConfig=function(){return this.config},t}()},function(t,e,n){"use strict";var r;n.d(e,"a",(function(){return s})),function(t){var e,n=0;t.getNext=function(){return void 0===e&&(e=function(){for(var t="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/",e=[],n=0;n<22;n++)e.push(t.charAt(Math.floor(Math.random()*t.length)));return e.join("")}()),new r(e,++n)},t.getNextChild=function(t){return new r(t.getString(),++t.nextChild)};var r=function(){function t(t,e){this.base=t,this.id=e,this.nextChild=0}return t.prototype.getString=function(){return this.base+"."+this.id},t}();t.CV=r}(r||(r={}));var i=n(1),o=function(t,e,n,r){return new(n||(n=Promise))((function(i,o){function a(t){try{s(r.next(t))}catch(t){o(t)}}function c(t){try{s(r.throw(t))}catch(t){o(t)}}function s(t){var e;t.done?i(t.value):(e=t.value,e instanceof n?e:new n((function(t){t(e)}))).then(a,c)}s((r=r.apply(t,e||[])).next())}))},a=function(t,e){var n,r,i,o,a={label:0,sent:function(){if(1&i[0])throw i[1];return i[1]},trys:[],ops:[]};return o={next:c(0),throw:c(1),return:c(2)},"function"==typeof Symbol&&(o[Symbol.iterator]=function(){return this}),o;function c(o){return function(c){return function(o){if(n)throw new TypeError("Generator is already executing.");for(;a;)try{if(n=1,r&&(i=2&o[0]?r.return:o[0]?r.throw||((i=r.return)&&i.call(r),0):r.next)&&!(i=i.call(r,o[1])).done)return i;switch(r=0,i&&(o=[2&o[0],i.value]),o[0]){case 0:case 1:i=o;break;case 4:return a.label++,{value:o[1],done:!1};case 5:a.label++,r=o[1],o=[0];continue;case 7:o=a.ops.pop(),a.trys.pop();continue;default:if(!(i=a.trys,(i=i.length>0&&i[i.length-1])||6!==o[0]&&2!==o[0])){a=0;continue}if(3===o[0]&&(!i||o[1]>i[0]&&o[1]<i[3])){a.label=o[1];break}if(6===o[0]&&a.label<i[1]){a.label=i[1],i=o;break}if(i&&a.label<i[2]){a.label=i[2],a.ops.push(o);break}i[2]&&a.ops.pop(),a.trys.pop();continue}o=e.call(t,a)}catch(t){o=[6,t],r=0}finally{n=i=0}if(5&o[0])throw o[1];return{value:o[0]?o[1]:void 0,done:!0}}([o,c])}}},c=function(){return 1e3*Date.now()};"object"==typeof window&&"object"==typeof window.performance&&"now"in window.performance&&(c=function(){return 1e3*Math.floor(window.performance.now())});var s=function(){function t(t,e,n){this._optionalEventFlags={},this._ended=!1,this._telemetryLogger=t,this._activityName=e,this._cv=n?r.getNextChild(n._cv):r.getNext(),this._dataFields=[],this._success=void 0,this._startTime=c()}return t.createNew=function(e,n){return new t(e,n)},t.prototype.createChildActivity=function(e){return new t(this._telemetryLogger,e,this)},t.prototype.setEventFlags=function(t){this._optionalEventFlags=t},t.prototype.addDataField=function(t){this._dataFields.push(t)},t.prototype.addDataFields=function(t){var e;(e=this._dataFields).push.apply(e,t)},t.prototype.setSuccess=function(t){this._success=t},t.prototype.setResult=function(t,e,n){this._result={code:t,type:e,tag:n}},t.prototype.endNow=function(){if(!this._ended){void 0===this._success&&void 0===this._result&&Object(i.d)(i.b.Warning,i.a.Core,(function(){return"Activity does not have success or result set"}));var t=c()-this._startTime;this._ended=!0;var e={duration:t,count:1,aggMode:0,cV:this._cv.getString(),success:this._success,result:this._result};return this._telemetryLogger.sendActivity(this._activityName,e,this._dataFields,this._optionalEventFlags)}Object(i.d)(i.b.Error,i.a.Core,(function(){return"Activity has already ended"}))},t.prototype.executeAsync=function(t){return o(this,void 0,void 0,(function(){var e=this;return a(this,(function(n){return[2,t(this).then((function(t){return e.endNow(),t})).catch((function(t){throw e.endNow(),t}))]}))}))},t.prototype.executeSync=function(t){try{var e=t(this);return this.endNow(),e}catch(t){throw this.endNow(),t}},t.prototype.executeChildActivityAsync=function(t,e){return o(this,void 0,void 0,(function(){return a(this,(function(n){return[2,this.createChildActivity(t).executeAsync(e)]}))}))},t.prototype.executeChildActivitySync=function(t,e){return this.createChildActivity(t).executeSync(e)},t}()},function(t,e,n){"use strict";n.d(e,"a",(function(){return r}));var r=function(){function t(){this._listeners=[]}return t.prototype.fireEvent=function(t){this._listeners.forEach((function(e){return e(t)}))},t.prototype.addListener=function(t){t&&this._listeners.push(t)},t.prototype.removeListener=function(t){this._listeners=this._listeners.filter((function(e){return e!==t}))},t.prototype.getListenerCount=function(){return this._listeners.length},t}()},function(t,e,n){"use strict";n.r(e);var r=n(6);n.d(e,"Contracts",(function(){return r.a}));var i=n(9);n.d(e,"ActivityScope",(function(){return i.a}));var o=n(2);n.d(e,"addContractField",(function(){return o.a}));var a=n(12);n.d(e,"getFieldsForContract",(function(){return a.a}));var c=n(4);n.d(e,"DataClassification",(function(){return c.a}));var s=n(13);for(var u in s)["default","Contracts","ActivityScope","addContractField","getFieldsForContract","DataClassification"].indexOf(u)<0&&function(t){n.d(e,t,(function(){return s[t]}))}(u);var d=n(0);n.d(e,"makeBooleanDataField",(function(){return d.a})),n.d(e,"makeInt64DataField",(function(){return d.d})),n.d(e,"makeDoubleDataField",(function(){return d.b})),n.d(e,"makeStringDataField",(function(){return d.e})),n.d(e,"makeGuidDataField",(function(){return d.c}));var f=n(3);n.d(e,"DataFieldType",(function(){return f.a}));var l=n(14);n.d(e,"getEffectiveEventFlags",(function(){return l.a}));var v=n(5);n.d(e,"SamplingPolicy",(function(){return v.e})),n.d(e,"PersistencePriority",(function(){return v.d})),n.d(e,"CostPriority",(function(){return v.a})),n.d(e,"DataCategories",(function(){return v.b})),n.d(e,"DiagnosticLevel",(function(){return v.c}));var p=n(15);for(var u in p)["default","Contracts","ActivityScope","addContractField","getFieldsForContract","DataClassification","makeBooleanDataField","makeInt64DataField","makeDoubleDataField","makeStringDataField","makeGuidDataField","DataFieldType","getEffectiveEventFlags","SamplingPolicy","PersistencePriority","CostPriority","DataCategories","DiagnosticLevel"].indexOf(u)<0&&function(t){n.d(e,t,(function(){return p[t]}))}(u);var y=n(1);n.d(e,"LogLevel",(function(){return y.b})),n.d(e,"Category",(function(){return y.a})),n.d(e,"onNotification",(function(){return y.e})),n.d(e,"logNotification",(function(){return y.d})),n.d(e,"logError",(function(){return y.c}));var h=n(8);n.d(e,"SuppressNexus",(function(){return h.b})),n.d(e,"SimpleTelemetryLogger",(function(){return h.a}));var g=n(16);n.d(e,"TelemetryLogger",(function(){return g.a}));var m=n(7);n.d(e,"cloneEvent",(function(){return m.a}));var b=n(17);for(var u in b)["default","Contracts","ActivityScope","addContractField","getFieldsForContract","DataClassification","makeBooleanDataField","makeInt64DataField","makeDoubleDataField","makeStringDataField","makeGuidDataField","DataFieldType","getEffectiveEventFlags","SamplingPolicy","PersistencePriority","CostPriority","DataCategories","DiagnosticLevel","LogLevel","Category","onNotification","logNotification","logError","SuppressNexus","SimpleTelemetryLogger","TelemetryLogger","cloneEvent"].indexOf(u)<0&&function(t){n.d(e,t,(function(){return b[t]}))}(u);var F=n(18);for(var u in F)["default","Contracts","ActivityScope","addContractField","getFieldsForContract","DataClassification","makeBooleanDataField","makeInt64DataField","makeDoubleDataField","makeStringDataField","makeGuidDataField","DataFieldType","getEffectiveEventFlags","SamplingPolicy","PersistencePriority","CostPriority","DataCategories","DiagnosticLevel","LogLevel","Category","onNotification","logNotification","logError","SuppressNexus","SimpleTelemetryLogger","TelemetryLogger","cloneEvent"].indexOf(u)<0&&function(t){n.d(e,t,(function(){return F[t]}))}(u)},function(t,e,n){"use strict";n.d(e,"a",(function(){return i}));var r=n(2);function i(t,e,n){var i=n.map((function(e){return{name:t+"."+e.name,value:e.value,dataType:e.dataType}}));return Object(r.a)(i,t,e),i}},function(t,e){},function(t,e,n){"use strict";n.d(e,"a",(function(){return o}));var r=n(5),i=n(1);function o(t){var e={costPriority:r.a.Normal,samplingPolicy:r.e.Measure,persistencePriority:r.d.Normal,dataCategories:r.b.NotSet,diagnosticLevel:r.c.FullEvent};return t.eventFlags&&t.eventFlags.dataCategories||Object(i.d)(i.b.Error,i.a.Core,(function(){return"Event is missing DataCategories event flag"})),t.eventFlags?(t.eventFlags.costPriority&&(e.costPriority=t.eventFlags.costPriority),t.eventFlags.samplingPolicy&&(e.samplingPolicy=t.eventFlags.samplingPolicy),t.eventFlags.persistencePriority&&(e.persistencePriority=t.eventFlags.persistencePriority),t.eventFlags.dataCategories&&(e.dataCategories=t.eventFlags.dataCategories),t.eventFlags.diagnosticLevel&&(e.diagnosticLevel=t.eventFlags.diagnosticLevel),e):e}},function(t,e){},function(t,e,n){"use strict";n.d(e,"a",(function(){return d}));var r,i=n(8),o=n(9),a=n(6),c=(r=function(t,e){return(r=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(t,e){t.__proto__=e}||function(t,e){for(var n in e)e.hasOwnProperty(n)&&(t[n]=e[n])})(t,e)},function(t,e){function n(){this.constructor=t}r(t,e),t.prototype=null===e?Object.create(e):(n.prototype=e.prototype,new n)}),s=function(t,e,n,r){return new(n||(n=Promise))((function(i,o){function a(t){try{s(r.next(t))}catch(t){o(t)}}function c(t){try{s(r.throw(t))}catch(t){o(t)}}function s(t){var e;t.done?i(t.value):(e=t.value,e instanceof n?e:new n((function(t){t(e)}))).then(a,c)}s((r=r.apply(t,e||[])).next())}))},u=function(t,e){var n,r,i,o,a={label:0,sent:function(){if(1&i[0])throw i[1];return i[1]},trys:[],ops:[]};return o={next:c(0),throw:c(1),return:c(2)},"function"==typeof Symbol&&(o[Symbol.iterator]=function(){return this}),o;function c(o){return function(c){return function(o){if(n)throw new TypeError("Generator is already executing.");for(;a;)try{if(n=1,r&&(i=2&o[0]?r.return:o[0]?r.throw||((i=r.return)&&i.call(r),0):r.next)&&!(i=i.call(r,o[1])).done)return i;switch(r=0,i&&(o=[2&o[0],i.value]),o[0]){case 0:case 1:i=o;break;case 4:return a.label++,{value:o[1],done:!1};case 5:a.label++,r=o[1],o=[0];continue;case 7:o=a.ops.pop(),a.trys.pop();continue;default:if(!(i=a.trys,(i=i.length>0&&i[i.length-1])||6!==o[0]&&2!==o[0])){a=0;continue}if(3===o[0]&&(!i||o[1]>i[0]&&o[1]<i[3])){a.label=o[1];break}if(6===o[0]&&a.label<i[1]){a.label=i[1],i=o;break}if(i&&a.label<i[2]){a.label=i[2],a.ops.push(o);break}i[2]&&a.ops.pop(),a.trys.pop();continue}o=e.call(t,a)}catch(t){o=[6,t],r=0}finally{n=i=0}if(5&o[0])throw o[1];return{value:o[0]?o[1]:void 0,done:!0}}([o,c])}}},d=function(t){function e(){return null!==t&&t.apply(this,arguments)||this}return c(e,t),e.prototype.executeActivityAsync=function(t,e){return s(this,void 0,void 0,(function(){return u(this,(function(n){return[2,this.createNewActivity(t).executeAsync(e)]}))}))},e.prototype.executeActivitySync=function(t,e){return this.createNewActivity(t).executeSync(e)},e.prototype.createNewActivity=function(t){return o.a.createNew(this,t)},e.prototype.sendActivity=function(t,e,n,r){return this.sendTelemetryEvent({eventName:t,eventContract:{name:a.a.Office.System.Activity.contractName,dataFields:a.a.Office.System.Activity.getFields(e)},dataFields:n,eventFlags:r})},e.prototype.sendError=function(t){var e=a.a.Office.System.Error.getFields("Error",t.error);return null!=t.dataFields&&e.push.apply(e,t.dataFields),this.sendTelemetryEvent({eventName:t.eventName,dataFields:e,eventFlags:t.eventFlags})},e}(i.a)},function(t,e){},function(t,e){},function(t,e,n){t.exports=n(11)}]);



OSFPerformance.officeExecuteEnd = OSFPerformance.now();
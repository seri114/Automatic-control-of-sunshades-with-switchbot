// src/trigger.ts

function setupTrigger(): string {
  var s = ""
  const sunsetDate = _fetchSunsetDate()
  s += `sunset time: ${sunsetDate.toLocaleTimeString()}\n`

  const closeTimeOffset = _getNumberProperty("closeTimeOffset", -3.5)
  s += _setTrigger("closeShade", _addHours(sunsetDate, closeTimeOffset))
  s += _setTrigger("setShadeAngle", _addHours(sunsetDate, closeTimeOffset, 2))
  s += _setTrigger("openShade", _addHours(sunsetDate, _getNumberProperty("openTimeOffset", -0.33)))
  return s
}

function checkIdDeviceIdValid(devicePropertyName: string, devices: any) {
  const deviceId = _getProperty(devicePropertyName)
  for (var i = 0; i < devices.length; i++) {
    var d = devices[i];
    if (d.deviceId === deviceId) {
      return
    }
  }
  throw new Error(`${devicePropertyName}  deviceId:${deviceId} not found in devices.`)
}

function checkIfValid() {
  var devices = _getDevices()
  checkIdDeviceIdValid("closeShadeDeviceId", devices)
  checkIdDeviceIdValid("openShadeDeviceId", devices)
  const setShadeAngleDeviceId = _getProperty("setShadeAngleDeviceId")
  if (setShadeAngleDeviceId !== "") {
    checkIdDeviceIdValid("setShadeAngleDeviceId", devices)
  }
}

function setupOffset(){
  var date = _fetchSunsetDate();
  var sunsetTime = `${date.getHours()}:${("00"+date.getMinutes()).slice(-2)}:${("00" + date.getSeconds()).slice(-2)}`
  _setProperty("SunsetTime", sunsetTime)
  const closeTimeOffset = _getNumberProperty("closeTimeOffset")
  const openTimeOffset = _getNumberProperty("openTimeOffset")
  return `SunsetTime: ${sunsetTime}\ncloseTimeOffset: ${closeTimeOffset}\nopenTimeOffset: ${openTimeOffset}`
}

function SetupAutomation() {
  checkIfValid()
  var s = setupOffset()
  _deleteTriggers("setupTrigger")
  ScriptApp.newTrigger("setupTrigger").
    timeBased().everyDays(1).atHour(9).create();
  s += setupTrigger()
  alert(s + "\nSunshade automation setup is now complete.")
}

function alert(msg: string) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(msg);
}

/**
 * Creates a Demo menu in Google Spreadsheets.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Setup Automation')
    .addItem('Setup', 'SetupAutomation')
    .addItem('Display devices', 'displayDevices')
    .addToUi();
};

function _getProperty(id: string, defaultString: string = ""): string {
  const namedRanges = SpreadsheetApp.getActive().getNamedRanges();
  for (var i = 0; i < namedRanges.length; i++) {
    if (namedRanges[i].getName() === id) {
      var a = namedRanges[i].getRange().getValue()
      console.log("found " + a.toString())
      return a;
    }
  }
  return defaultString;
}

function _getNumberProperty(id: string, defaultNumber: number = 0): number {
  var v = _getProperty(id, undefined)
  if (!v) {
    return defaultNumber;
  }
  return parseFloat(v)
}

function _setProperty(id: string, param: any){
  const namedRanges = SpreadsheetApp.getActive().getNamedRanges();
  for (var i = 0; i < namedRanges.length; i++) {
    if (namedRanges[i].getName() === id) {
      namedRanges[i].getRange().setValue(param)
    }
  }
}


function closeShade() {
  _pressSwitchbot(_getProperty("closeShadeDeviceId", "set valid closeShadeDeviceId"))
}

function setShadeAngle() {
  const setShadeAngleDeviceId = _getProperty("setShadeAngleDeviceId")
  if (setShadeAngleDeviceId === "") {
    console.log("skip set shade angle")
  }
  _pressSwitchbot(setShadeAngleDeviceId)
}
function openShade() {
  _pressSwitchbot(_getProperty("openShadeDeviceId", "set valid openShadeDeviceId"))
}

function displayDevices() {
  alert(JSON.stringify(_getDevices(), null, 2));
}

function _getRequestHeaders(): GoogleAppsScript.URL_Fetch.HttpHeaders {
  var token = getSwitchbotToken()
  var requestHeaders = {
    'Authorization': token,
    "Content-type": "application/json; charset=utf-8",
  };
  return requestHeaders
}

function _getRequestOptions(method: GoogleAppsScript.URL_Fetch.HttpMethod, payload: GoogleAppsScript.URL_Fetch.Payload | undefined = undefined): GoogleAppsScript.URL_Fetch.URLFetchRequestOptions {
  var requestOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    "headers": _getRequestHeaders(),
    "method": method,
    "payload": payload
  }
  return requestOptions
};

function _fetchSwitchbotCommand(api: string, requestOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions | undefined = undefined): any {
  try {
    const API_URL = "https://api.switch-bot.com/v1.0/"
    var url_list = API_URL + api;
    var response = UrlFetchApp.fetch(url_list, requestOptions ?? _getRequestOptions('get'));
    var responseText = response.getContentText();
    var a = JSON.parse(responseText)
    return a
  } catch (e) {
    throw new Error(`Check if switchbotToken is valid: ${getSwitchbotToken()}\n\n${e}`)
  }
}

function _getDevices(): any {
  return _fetchSwitchbotCommand("devices").body.deviceList
}

function getSwitchbotToken(): string {
  const token = _getProperty("switchbotToken")
  if (!token) {
    throw new Error("Input Valid SwitchbotToken")
  }
  return token
}

function _pressSwitchbot(deviceId: string) {
  var payload =
  {
    "command": "press",
    "parameter": "default",
    "commandType": "command"
  };
  var requestOptions = _getRequestOptions('post', JSON.stringify(payload))
  var ret = _fetchSwitchbotCommand(`devices/${deviceId}/commands`, requestOptions)
  console.log(`pressSwitchbot ${deviceId} ${JSON.stringify(ret)}`)
}

function _fetchSunsetDate(): Date {
  var lat = _getNumberProperty("lat", 35.635032)
  var lng = _getNumberProperty("lng", 139.756410)
  var requestUrl = `https://api.sunrise-sunset.org/json?&formatted=0&lat=${lat}&lng=${lng}`
  var response = UrlFetchApp.fetch(requestUrl)
  var responseText = response.getContentText()
  var parsedResponse = JSON.parse(responseText)
  var sunsetTime = parsedResponse.results.sunset
  var s = new Date(Date.parse(sunsetTime))
  var d = new Date()
  d.setHours(s.getHours(), s.getMinutes())
  return d
}

function _addHours(date: Date, numOfHours: number, numOfMinutes?: number): Date {
  return new Date(date.getTime() + (numOfHours * 60 + (numOfMinutes ?? 0)) * 60 * 1000)
}


function _deleteTriggers(functionName: string) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() !== functionName) return;
    ScriptApp.deleteTrigger(trigger);
  });
}

function _setTrigger(functionName: string, date: Date): string {
  var s = `${functionName} at ${date.toLocaleTimeString()}\n`
  console.log(`set ${functionName} at ${date.toLocaleDateString()} ${date.toLocaleTimeString()}`)
  _deleteTriggers(functionName)
  if (Date.now() > date.getTime()) {
    console.log("error: old trigger")
    return s
  }
  ScriptApp.newTrigger(functionName).
    timeBased().
    at(date).
    create();
  return s
}
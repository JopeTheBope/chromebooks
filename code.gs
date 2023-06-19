//Authentication as service account
function createService(userToImpersonate) {
  return OAuth2.createService('ChromeOSDevices')
    .setSubject(userToImpersonate)
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')
    .setPrivateKey(serviceAccount.private_key)
    .setIssuer(serviceAccount.client_email)
    .setScope(['https://www.googleapis.com/auth/admin.directory.device.chromeos.readonly'])
}

function chromeos() {
  //Authentication as service account
  let oauthService = createService('admin email')

  //Google Admin API URL and CustomerID
  let url = `https://admin.googleapis.com/admin/directory/v1/customer/my_customer/devices/chromeos`

  //Authorization
  let headers = {Authorization: `Bearer ${oauthService.getAccessToken()}`}

  //Loop API request through pages of 'chromeosdevices.list'
  let currentPageToken = null
  do{
    let options = {method: 'get', headers, contentType: 'application/json', muteHttpExceptions: true, pageToken: currentPageToken}
    let page = JSON.parse(UrlFetchApp.fetch(url,options).getContentText())
    let devices=page.chromeosdevices
    let devicesToUpdate = []
    let devicesToAdd = []
    for (let i = 0; i<=devices.length; i++) {
      let device = devices[i]
      if (device?.annotatedAssetId !== undefined && device?.deviceId !== undefined) {
        let response = appsheetRead(device)
        Logger.log(response)
        //If device record is in database, update it. Else, add it.
        if (response[0]) {
          devicesToUpdate = pushDevicesToUpdate(device,devicesToUpdate)
          Logger.log("Update: "+devicesToUpdate.length)
        } else {
          devicesToAdd = pushDevicesToAdd(device,devicesToAdd)
          Logger.log("Add: "+devicesToAdd.length)
        }
      }
    }
    if (devicesToUpdate.length > 0) {appsheetUpdate(devicesToUpdate)}
    if (devicesToAdd.length > 0) {appsheetAdd(devicesToAdd)}
    if (currentPageToken !== page.nextPageToken) {
      currentPageToken=page.nextPageToken
    } else {
      currentPageToken = null
    }
  } while (currentPageToken)
}

function pushDevicesToUpdate(device,devicesToUpdate) {
  devicesToUpdate.push ({
    "DeviceName": device.annotatedAssetId,
    "UUID": device.deviceId,
    "LastSyncTime": Utilities.formatDate(new Date(), "Europe/London", "dd/MM/yyyy HH:mm:ss"),
    "LastSyncUser": device.recentUsers[0].email,
    "OsVersion": device.osVersion,
    "Ram": device.systemRamTotal,
    "StorageType": device.storageType,
    "StorageSize": device.diskVolumeReports[0].volumeInfo[0].storageTotal,
    "Firmware": device.firmwareVersion,
    "ConnectedNetwork": device.connectedNetwork,
    "EthernetMacAddress": device.ethernetMacAddress,
    "WifiMacAddress": device.macAddress
    })
  return devicesToUpdate
}

function pushDevicesToAdd(device,devicesToAdd) {
  devicesToAdd.push({
    "DeviceName": device?.annotatedAssetId,
    "Model":device?.model,
    "SerialNumber": device?.serialNumber,
    "UUID": device?.deviceId,
    "LastSyncTime": Utilities.formatDate(new Date(), "Europe/London", "dd/MM/yyyy HH:mm:ss"),
    "LastSyncUser": device?.recentUsers?.[0].email,
    "OperatingSystem": "ChromeOS",
    "OsVersion": device?.osVersion,
    "Processor": device?.cpuInfo?.model,
    "Ram": device?.systemRamTotal,
    "StorageType": device?.storageType,
    "StorageSize": device?.diskVolumeReports[0].volumeInfo[0].storageTotal,
    "Firmware": device?.firmwareVersion,
    "ConnectedNetwork": device?.connectedNetwork,
    "EthernetMacAddress": device?.ethernetMacAddress,
    "WifiMacAddress": device?.macAddress
    })
  return devicesToAdd
}

//Find device in database
function appsheetRead(device) {
  let payload = JSON.stringify({
    "Action": "Find",
    "Properties": {
      "Locale": "en-GB",
      "Timezone": "GMT Standard Time",
      "Selector": `Filter(Devices,[UUID]='${device?.deviceId}')`
    },
    "Rows": []
  })
  let response = sendRequest(payload)
  return response
}

//Update device record in database
function appsheetUpdate(devicesToUpdate) {
  let payload = JSON.stringify({
    "Action": "Edit",
    "Properties": {
      "Locale": "en-GB",
      "Timezone": "GMT Standard Time"
    },
    "Rows": devicesToUpdate
  })
  let response = sendRequest(payload)
}

//Add device to database
function appsheetAdd(devicesToAdd) {
  let payload = JSON.stringify({
    "Action": "Add",
    "Properties": {
      "Locale": "en-GB",
      "Timezone": "GMT Standard Time"
    },
    "Rows": devicesToAdd
  })
  let options = {method: 'post', headers: {"ApplicationAccessKey": "", "Content-Type": "application/json"}, payload: payload}
  let response = UrlFetchApp.fetch("https://api.appsheet.com/api/v2/apps/6c2efb23-6c3d-4e79-90f6-f65c6ef40463/tables/Devices/Action?", options)
}

function sendRequest(payload) {
  let options = {method: 'post', headers: {"ApplicationAccessKey": "", "Content-Type": "application/json"}, payload: payload}
  let response = JSON.parse(UrlFetchApp.fetch("https://api.appsheet.com/api/v2/apps/6c2efb23-6c3d-4e79-90f6-f65c6ef40463/tables/Devices/Action?", options).getContentText())
  return response
}

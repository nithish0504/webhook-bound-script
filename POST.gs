function myFunction() {
  
}

const documentProperties = PropertiesService.getDocumentProperties();
let ok200Status = false;
let logTimeStamp = false;


function _getBaseUrl() {
  return 'https://api.runo.in/qa';
}

function onOpen(e) {
  if (documentProperties.getProperty('Authorized') !== 'true') {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Runo Webhooks')
      .addItem('Authorize Webhook', 'authorizeScript')
      .addToUi();
  }
}

function authorizeScript() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Authorization successful.', "🪝 Runo Webhooks");
  documentProperties.setProperty('Authorized', 'true');
}

function restructureParameters(parameters){
  let queryParams = {}
  for(let param in parameters){
    queryParams[param]=parameters[param][0]
  }
  return queryParams
}

function getProcess(apiKey){
  const url = `${_getBaseUrl()}/process`;
    const options = {
      'method': 'get',
      'contentType': 'application/json',
      'headers': {
        'Auth-Key': apiKey
      }
    };
    const response = JSON.parse(UrlFetchApp.fetch(url, options));
    if (response['statusCode'] == 0) {
      Logger.log('Successfully fetched the process');
      return {
        isSucess:true,
        data:response.data
      }
    }else{
      return {
        isSucess:false,
        data:'',
        error:response.message
      }
    }
}

function getAllActiveSheets(){
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = activeSpreadsheet.getSheets();
  const sheetsArray = allSheets.map((_sheet)=>{
    return{
      sheetName:_sheet.getSheetName(),
      sheetId:_sheet.getSheetId().toString()
    }
  })
  return sheetsArray
}

function getAllActiveSheetsObj(){
  let sheetsArray = getAllActiveSheets()
  let sheetsObj = {}
  sheetsArray.forEach((_sheet)=>{
    sheetsObj[_sheet.sheetId]=_sheet.sheetName
  })
  return sheetsObj
}

function getWebhookConfig(processId,input_type){
  let storedWebhookData = documentProperties.getProperty("webhookData")
  if(storedWebhookData){
    storedWebhookData = JSON.parse(storedWebhookData)
    if(storedWebhookData[processId]){
      if(storedWebhookData[processId][input_type]){
        return {
          isSuccess:true,
          data:storedWebhookData[processId][input_type]
        }
      }
    }
  }
  return{
    isSuccess:false,
    data:''
  }
}

const callLogFieldMapping = {
  //"Column_Name_In_Excel": "Field_Name_In_Data"
  "callId":"callId",
  "Caller Id": "callerId",
  "Caller Name": "calledBy",
  "Name": "name",
  "Phone": "phoneNumber",
  "Call Start Date": "startTime",
  "Duration (in Sec)": "duration",
  "Call Type": "type",
  "Call Tag": "tag",
  "Status": "status",
  "Recording Path":"recordingPath"
};


function postCallLog(body,queryParams){
  let webhookConfigResponse = getWebhookConfig(body.processId,"callLogs")
  if(!webhookConfigResponse.isSuccess){
    return {
      statusCode:"-1",
      status:"error",
      error:"Webhook config object missing"
    }
  }
  let webhookConfig = webhookConfigResponse.data

  const variables = webhookConfig;
  const apiKey = variables.apiKey;
  let targetSheetName = '';
  let targetSheetId = variables["trackingSheetName"];
  const fieldMapping  = variables["fieldMapping"] ? variables["fieldMapping"] : callLogFieldMapping;
  const dateCustomFields = variables.date_columns;
  let processId = body.processId;
  let process_Name = variables["defaultProcessName"] ? variables["defaultProcessName"] : '';
  let process_data=[]
  let columns = [];

  if (!apiKey) {
    Logger.log('Parameter apiKey is mandatory. Please provide proper value');
    return{
      status: 'error',
      message: 'Api key Missing in Webhook data'
    }
  }

//   if(!process_Name){
//     let fetchProcessResponse = getProcess(apiKey);
//     if(fetchProcessResponse.isSucess){
//       process_data=fetchProcessResponse.data
//     }
//   }
  console.log(process_data)

  if(targetSheetName === ''){
    let sheetsObj = getAllActiveSheetsObj()
    if(targetSheetId && targetSheetId in sheetsObj){
      targetSheetName = sheetsObj[targetSheetId]
    }else{
      targetSheetName = `${process_Name}_callLogs`
    }
  }

  for(key in fieldMapping){
    columns.push(key)
  }

  let sheet = createSheet_(SpreadsheetApp.getActive(), targetSheetName);
  if (!sheet) {
    Logger.log('Given interactions sheet with name ' + targetSheetName + ' is not found. Please correct the sheet name');
    return{
      status: 'error',
      message: `Given target sheet with name ${targetSheetName} is not found. Please correct the sheet name`
    }
  }

  let lastRowIndex = sheet.getLastRow();

  if(body.recordingPath){
    columns.push("Recording Path");
  }

  if (lastRowIndex === 0.0) {
    Logger.log("Column names are missing in sheet. Adding the column names")
    sheet.appendRow(columns)
    lastRowIndex = sheet.getLastRow()
    sheet.setFrozenRows(1);
  }
  const headerValues = sheet.getDataRange().getValues()[0];
  Logger.log(headerValues)

  let sheetDataToBeWritten = []

  const getValue = (o, path) => {
    const props = path.split(".");
    let value = null;
    if (props.length > 1) {
      props.forEach((_prop) => {
        value = (value == null) ? o[_prop] : value[_prop]
      })
    } else {
      value = o[path]
    }
    return value || "";
  }

  let jsonData = Array.isArray(body) ? body : [body];

  jsonData.forEach((_interaction) => {
    let rowData = [];
    headerValues.forEach((_header) => {
      let value = "";
      let path = _header;
      if (fieldMapping.hasOwnProperty(_header)) {
        path = fieldMapping[_header];
      }
      value = getValue(_interaction, path);
      if (path === "createdAt" || path === "startTime") {
        value = new Date(value * 1000).toLocaleString('en-GB', { timeZone: 'Asia/Kolkata' });
      } else {
        value = value.toString();
      }
      if (value != null) {
        rowData.push(value);
      }
    })
    sheetDataToBeWritten.push(rowData);
  })

  if (sheetDataToBeWritten.length > 0) {
    writeToSheet_(sheet, lastRowIndex + 1, 1, sheetDataToBeWritten.length, Object.getOwnPropertyNames(headerValues).length - 1,sheetDataToBeWritten);
    Logger.log(`Successfully updated the sheet with callLogs`);
    return{
      status: 'success',
      message: 'Data logged successfully'
    };
  } else {
    Logger.log(`No callLogs are found in the data`);
    return{
      status: 'success',
      message: 'No callLogs are found in the data'
    };
  }
}

function postInteraction(body,queryParams){
  let webhookConfigResponse = getWebhookConfig(body.processId,"interactions")
  if(!webhookConfigResponse.isSuccess){
    return {
      statusCode:"-1",
      status:"error",
      error:"Webhook config object missing"
    }
  }
  let webhookConfig = webhookConfigResponse.data

  const variables = webhookConfig;
  const apiKey = variables.apiKey;
  let targetSheetName = '';
  let targetSheetId = variables["trackingSheetName"]
  const fixedFieldMapping = variables.fixedFieldMapping;
  const customFieldMapping = variables.customFieldMapping;
  const dateCustomFields = variables.date_columns;
  let processId = body.processId
  let columns = [];
  let process_data=[]
  let process_Name = variables["defaultProcessName"] ? variables["defaultProcessName"] : ''

  if (!apiKey) {
    Logger.log('Parameter apiKey is mandatory. Please provide proper value');
    return{
      status: 'error',
      message: 'Api key Missing in Webhook data'
    }
  }

//   if(!process_Name){
//     let fetchProcessResponse = getProcess(apiKey);
//     if(fetchProcessResponse.isSucess){
//       process_data=fetchProcessResponse.data
//     }
//   }
  console.log(process_data)

  if(targetSheetName === ''){
    let sheetsObj = getAllActiveSheetsObj()
    if(targetSheetId && targetSheetId in sheetsObj){
      targetSheetName = sheetsObj[targetSheetId]
    }else{
      targetSheetName = `${process_Name}_interactions`
    }
  }

  let interaction_fields={
    'Interaction Date': 'createdAt',
    'Process Name':'processName'
  }
  for (let i in interaction_fields) {
    columns.push(i)
    fixedFieldMapping[i] = interaction_fields[i]
  }
  for (let i in fixedFieldMapping) {
    columns.push(i)
  }
  for (let i in customFieldMapping) {
    columns.push(i)
  }

  let sheet = createSheet_(SpreadsheetApp.getActive(), targetSheetName);
  if (!sheet) {
    Logger.log('Given interactions sheet with name ' + targetSheetName + ' is not found. Please correct the sheet name');
    return{
      status: 'error',
      message: `Given target sheet with name ${targetSheetName} is not found. Please correct the sheet name`
    }
  }

  let lastRowIndex = sheet.getLastRow();
  if (lastRowIndex === 0.0) {
    Logger.log('Column names are missing in sheet. Adding the column names')
    sheet.appendRow(columns)
    lastRowIndex = sheet.getLastRow()
    sheet.setFrozenRows(1);
  }
  const headerValues = sheet.getDataRange().getValues()[0];
  Logger.log(headerValues)

  let jsonData = Array.isArray(body) ? body : [body];
  let sheetDataToBeWritten = []

  const getValue = (o, path) => {
    const props = path.split('.');
    let value = null;
    if (props.length > 1) {
      props.forEach((_prop) => {
        value = (value == null) ? o[_prop] : value[_prop]
      })
    } else {
      value = o[path]
    }
    return value || '';
  }
  let keyValueMap = {}

  jsonData.forEach((_interaction) => {
    let rowData = []
    headerValues.forEach((_header) => {
      let value = ''
      if (fixedFieldMapping.hasOwnProperty(_header)) {
        let path = fixedFieldMapping[_header];
        value = getValue(_interaction, path);
        if (path === 'createdAt') {
          value = new Date(value * 1000).toLocaleString('en-GB', { timeZone: 'Asia/Kolkata' });
        } else if (path === 'location') {
          value = value['lat'] ? `https://google.com/maps?q=${value['lat']},${value['long']}` : '';
        } else if (path === 'priority') {
          switch (value) {
            case 3:
              value = 'High';
              break;
            case 2:
              value = 'Medium';
              break;
            case 1:
              value = 'Low';
              break;
            default:
              value = 'Not Set';
              break;
          }
        } else if (path !== 'customer.company.address.pincode') {
          value = value.toString();
        }
        keyValueMap[path]=value
      } else {
        let mappedHeader = _header;
        if (customFieldMapping.hasOwnProperty(_header)) {
          mappedHeader = customFieldMapping[_header];
        }
        let fieldIndex = _interaction.userFields.findIndex((_x) => _x.name.toLowerCase() === mappedHeader.toLowerCase());
        if (fieldIndex !== -1) {
          value = _interaction.userFields[fieldIndex].value;
          if (dateCustomFields.indexOf(_header) !== -1) {
            value = new Date(value * 1000).toLocaleString('en-GB', { timeZone: 'Asia/Kolkata' });
          }
        }
      }
      if(_header=='Process Name'){
          value = process_Name
      }
      if (value !== null) {
        rowData.push(value);
      }
      keyValueMap[_header]=value
    })
    sheetDataToBeWritten.push(rowData);
  });

  if (sheetDataToBeWritten.length > 0) {
    writeToSheet_(sheet, lastRowIndex + 1, 1, sheetDataToBeWritten.length, Object.getOwnPropertyNames(headerValues).length - 1,sheetDataToBeWritten);
    Logger.log(`Successfully updated the sheet with interactions`);
    return{
      status: 'success',
      message: 'Data logged successfully',
      keyValueMap:keyValueMap,
      data:sheetDataToBeWritten,
      name:process_Name
    }
  } else {
    Logger.log(`No interactions are found in the data`);
    return{
      status: 'success',
      message: 'No interactions are found in the data'
    };
  }

}

function doGet(e) {
  let params = e.parameters;

  if(params.data_type[0]==="variables"){
    let response = {
      status:"success",
      data:JSON.parse(documentProperties.getProperty("webhookData"))
    }
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
  }
  let response = {
    status:"error",
    message:"please send a valid request"
  }
}

function doPost(e) {

  let { parameters, postData: { contents, type } = {} } = e;
  let response = {};
  console.log(e)

  if (type === 'text/plain' || type === 'text/html' || type === 'application/xml') {
    response = {
      status: 'error',
      message: `Unsupported data-type: ${type}`
    }
    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
  }

  if (type === 'application/json' || (type === '' && contents.length > 0)) {
    try {
      jsonData = JSON.parse(contents);
    } catch (e) {
      response = {
        status: 'error',
        message: 'Invalid JSON format'
      };
      return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
    };
  }

  let queryParams = restructureParameters(parameters)
  let postBody = JSON.parse(contents)

  if(queryParams.invocation_type && queryParams.invocation_type==="initiaite_variables"){
    let webhookData = JSON.parse(contents);
    let storedWebhookData = documentProperties.getProperty("webhookData");
    if(storedWebhookData){
      storedWebhookData = JSON.parse(storedWebhookData)
      for(processId in webhookData){
        if(processId in storedWebhookData){
          for(input_type in webhookData[processId]){
            storedWebhookData[processId][input_type]=webhookData[processId][input_type]
          }
        }else{
          storedWebhookData[processId] = webhookData[processId]
        }
      }
    }else{
      storedWebhookData = webhookData
    }
    documentProperties.setProperty("webhookData",JSON.stringify(storedWebhookData))
    response = {
      statusCode:"0",
      status: 'success',
      message: 'Variables Initiated successfully'
    }
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
  }

  if(queryParams.data_type && queryParams.data_type!==""){
    switch (queryParams.data_type){
      case "interactions":
        let interactionResponse = postInteraction(postBody,queryParams)
        return ContentService.createTextOutput(JSON.stringify(interactionResponse)).setMimeType(ContentService.MimeType.JSON);
      case "callLogs":
        let callLogresponse = postCallLog(postBody,queryParams)
        return ContentService.createTextOutput(JSON.stringify(callLogresponse)).setMimeType(ContentService.MimeType.JSON);
      default:
        let defaultResponse = {
          statusCode:"-1",
          status:"error",
          message:"please send a valid data_type"
        }
        return ContentService.createTextOutput(JSON.stringify(defaultResponse)).setMimeType(ContentService.MimeType.JSON);
    }
  }else{
    if("callerId" in postBody){
      let callLogresponse = postCallLog(postBody,queryParams)
      return ContentService.createTextOutput(JSON.stringify(callLogresponse)).setMimeType(ContentService.MimeType.JSON);
    }else{
      let interactionResponse = postInteraction(postBody,queryParams)
      return ContentService.createTextOutput(JSON.stringify(interactionResponse)).setMimeType(ContentService.MimeType.JSON);
    }
  }

  let invalidResponse = {
    statusCode:"-1",
    status:"error",
    message:"please send a valid request"
  }
  return ContentService.createTextOutput(JSON.stringify(invalidResponse)).setMimeType(ContentService.MimeType.JSON);

}

function writeToSheet_(sheet, startRow, startColumn, totalRows, totalColumns, data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  sheet.getRange(startRow, startColumn, totalRows, totalColumns).setValues(data);
  SpreadsheetApp.flush();
  lock.releaseLock();
}

function createSheet_(activeSpreadsheet, sheetName) {
  let sheet = activeSpreadsheet.getSheetByName(sheetName);
  if (sheet == null) {
    sheet = activeSpreadsheet.insertSheet();
    sheet.setName(sheetName);
  }

  return sheet;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var myMenu = ui.createMenu('Strobbo');
  myMenu.addItem('Import employees', 'uploadEmployees');
  myMenu.addItem('Clear table', 'clearTable');
  myMenu.addItem('Clear result', 'clearResult');
  myMenu.addToUi();
}

/**
 * Posts a json object and returns ok on success and the response body on error.
 *
 * @param {string} url
 * @param {string} apiKey The API key to use. it's placed in the headers with the name "ApiKey"
 * @param {object} body The body to post
 * @returns {string} 'ok' on a successful upload, the result body otherwise
 */
function PostJSONApiKey(url, apiKey, body) {
  const fetchResult = UrlFetchApp.fetch(url, {
    method: 'POST',
    payload: JSON.stringify(body),
    headers: {ApiKey: apiKey, 'Content-Type': 'application/json'},
    muteHttpExceptions: true,
  });
  if (fetchResult.getResponseCode() >= 200 && fetchResult.getResponseCode() < 300) return 'ok';
  return fetchResult.getContentText();
}

/**
 * @typedef {Object} dtoField Field in the dto array to validate input
 * @prop {string} name
 * @prop {boolean} [required] Defaults to false
 * @prop {'string' | 'date' | 'number'} [type] Defaults to 'string'
 */

/** @type {dtoField[]} */
const CreateOrUpdateEmployeeDto = [
  { name: 'employeeNumber', required: true },
  { name: 'nationalInsuranceNumber' },
  { name: 'firstName', required: true },
  { name: 'lastName', required: true },
  { name: 'displayName' },
  { name: 'email', required: true },
  { name: 'workspaces' },
  { name: 'contractStartDate', type: 'date' },
  { name: 'contractEndDate', type: 'date' },
  { name: 'externalId', type: 'number' },
  { name: 'externalNumber' },
  { name: 'street' },
  { name: 'houseNumber' },
  { name: 'city' },
  { name: 'postalCode' },
  { name: 'country' },
  { name: 'nationality' },
  { name: 'dateOfBirth', type: 'date' },
  { name: 'placeOfBirth' },
  { name: 'countryOfBirth' },
  { name: 'gender' },
  { name: 'maritalStatus' },
  { name: 'maritalStatusDate', type: 'date' },
  { name: 'iban' },
  { name: 'bic' },
  { name: 'costPerHour', type: 'number' },
  { name: 'wagePerHour', type: 'number' },
  { name: 'phoneNumber' },
  { name: 'mobileNumber' },
  { name: 'language' },
  { name: 'functionTitle' },
  { name: 'hoursPerWeek', type: 'number' },
];
const genderEnum = [
  [0, 'Unknown'],
  [1, 'Male'],
  [2, 'Female'],
];
const maritalStatusEnum = [
  [0, 'Unknown'],
  [1, 'Unmarried'],
  [2, 'Married'],
  [3, 'Legally living together'],
  [4, 'Domestic living together'],
  [5, 'Legally divorced'],
  [6, 'Domestic divorced'],
  [7, 'Widowed'],
];
const convertEnum = (input, enums) => {
  for (const [enumValue, enumName] of enums) {
    if (input === enumName || input.toString() === enumValue.toString()) return enumValue;
  }
  throw new TypeError(`Enum value ${input} not found.`);
};

const ignoreInUpload = ['uploadStatus'];

function uploadEmployees() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  if (sheet.getName() !== 'Import') {
    ui.alert('Import function should be called on the sheet named "import"');
    return;
  }
  const data = sheet.getDataRange().getValues();

  // ui.alert(JSON.stringify(data, undefined, 2));

  if (data.length < 2) {
    ui.alert('No rows detected to upload');
    return;
  }

  const headers = data[0];
  const headersNotFound = headers.filter(
    (header) => !ignoreInUpload.includes(header) && !CreateOrUpdateEmployeeDto.find((field) => field.name === header),
  );
  if (headersNotFound.length) {
    ui.alert(`The following fields were not recognized: ${headersNotFound.join(', ')}`);
    return;
  }

  const errors = [];
  const objectsToPost = [];

  for (let iRow = 1; iRow < data.length; iRow++) {
    const dataRow = data[iRow];
    const objectToPost = {};

    for (let iCol = 0; iCol < dataRow.length; iCol++) {
      const cellData = dataRow[iCol];
      const currentHeader = headers[iCol];
      if (ignoreInUpload.includes(currentHeader)) continue;
      const currentField = CreateOrUpdateEmployeeDto.find((field) => field.name === currentHeader);
      if (!currentField) {
        errors.push(`Field ${currentField} not recognized`);
        continue;
      }

      if (cellData === '') {
        if (currentField.required) {
          errors.push(`${currentField.name} is required, but is missing a value on row ${iRow + 1}`);
        } else {
          continue;
        }
      }

      switch (currentField.type) {
        case undefined:
        case 'string':
          objectToPost[currentHeader] = cellData.toString();
          break;
        case 'number':
          if (typeof cellData !== 'number') {
            errors.push(
              `${currentField.name} must be a number, but a ${typeof cellData} was detected on row ${iRow + 1}`,
            );
          }
          objectToPost[currentHeader] = cellData;
          break;
        case 'date':
          if (!(cellData instanceof Date)) {
            errors.push(
              `${currentField.name} must be a date, but a ${typeof cellData} was detected on row ${iRow + 1}`,
            );
          }
          objectToPost[currentHeader] = cellData;
          break;
      }
    }

    // custom

    for (const addressField of ['street', 'houseNumber', 'city', 'postalCode', 'country']) {
      if (!(addressField in objectToPost)) continue;
      if (!('address' in objectToPost)) objectToPost.address = {};
      objectToPost.address[addressField] = objectToPost[addressField];
      delete objectToPost[addressField];
    }

    if ('gender' in objectToPost) {
      try {
        objectToPost.gender = convertEnum(objectToPost.gender, genderEnum);
      } catch (e) {
        if (e instanceof TypeError) {
          errors.push(
            `Enum value for "${objectToPost.gender}" for field gender not found, possible values are ` +
              `"${genderEnum.flat().join('" ,"')}" on row ${iRow + 1}`,
          );
        } else {
          throw e;
        }
      }
    }

    if ('maritalStatus' in objectToPost) {
      try {
        objectToPost.maritalStatus = convertEnum(objectToPost.maritalStatus, maritalStatusEnum);
      } catch (e) {
        if (e instanceof TypeError) {
          errors.push(
            `Enum value for "${objectToPost.gender}" for field maritalStatusEnum not found, ` +
              `possible values are "${genderEnum.flat().join('" ,"')} on row ${iRow + 1}"`,
          );
        } else {
          throw e;
        }
      }
    }

    // google sheets threads date as iso string, meaning the time zone offset will push the date a day early
    if ('dateOfBirth' in objectToPost) {
      objectToPost.dateOfBirth.setHours(objectToPost.dateOfBirth.getHours() + 2);
      objectToPost.dateOfBirth = objectToPost.dateOfBirth.toISOString().slice(0, 10);
    }
    if ('maritalStatusDate' in objectToPost) {
      objectToPost.maritalStatusDate.setHours(objectToPost.maritalStatusDate.getHours() + 2);
      objectToPost.maritalStatusDate = objectToPost.maritalStatusDate.toISOString().slice(0, 10);
    }

    if (objectToPost.workspaces) {
      objectToPost.workspaces = objectToPost.workspaces.split(',').map((workspace) => workspace.trim());
    } else {
      objectToPost.workspaces = [];
    }

    if (!objectToPost.externalId && !objectToPost.externalNumber) {
      errors.push(`An employee must have either an externalId or an externalNumber on row ${iRow + 1}"`);
    }

    objectsToPost.push(objectToPost);
  }

  if (errors.length) {
    ui.alert(`Input validation failed:\n - ${errors.join('\n - ')}`);
    return;
  }

  // ui.alert(JSON.stringify(objectsToPost, undefined, 2));

  const file = SpreadsheetApp.getActive();

  const setupSheet = file.getSheetByName('Setup');
  if (!setupSheet) throw new Error('Setup sheet not found');
  const rowWithApiKey = setupSheet
    .getDataRange()
    .getValues()
    .find((row) => row[0] === 'Api Key');
  if (!rowWithApiKey) throw new Error('Row with api key in setup sheet not found');
  const apiKey = rowWithApiKey[1];

  const apiSheet = file.getSheetByName('API');
  if (!apiSheet) throw new Error('API sheet not found');
  const rowApiEndpoint = apiSheet
    .getDataRange()
    .getValues()
    .find((row) => row[0] === 'POST Employees');
  if (!rowApiEndpoint) throw new Error('Row with api endpoint "POST Employees" in setup sheet not found');
  const endpoint = rowApiEndpoint[1];

  const uploadErrors = [];
  for (let i = 0; i < objectsToPost.length; i++) {
    const result = PostJSONApiKey(endpoint, apiKey, objectsToPost[i]);
    const resultCell = sheet.getRange(i + 2, 1);
    resultCell.setValue(result);
    if (result === 'ok') {
      resultCell.setFontColor('green');
    } else {
      resultCell.setFontColor('red');
      uploadErrors.push([i + 2, result]);
    }
  }

  if (uploadErrors.length === 0) {
    ui.alert('Upload successful');
  } else {
    ui.alert(
      `There were errors while uploading:\n${uploadErrors
        .map(([rowNr, error]) => `- row ${rowNr}: ${error}`)
        .join('\n')}`,
    );
  }
}

function clearTable() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== 'Import') {
    SpreadsheetApp.getUi().alert('This function is meant to be called on the sheet named "import"');
    return;
  }

  // index starts at 1
  const rangeToClear = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  // ui.alert(rangeToClear.getValues());
  rangeToClear.clearContent();
}

function clearResult() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== 'Import') {
    SpreadsheetApp.getUi().alert('This function is meant to be called on the sheet named "import"');
    return;
  }

  // index starts at 1
  const rangeToClear = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  // SpreadsheetApp.getUi().alert(rangeToClear.getValues());
  rangeToClear.clear();
}

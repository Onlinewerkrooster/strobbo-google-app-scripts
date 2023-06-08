function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const myMenu = ui.createMenu('Strobbo');
  myMenu.addItem('Import revenue', 'importRevenue');
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

const toTwoDigits = (number) => number.toString().padStart(2, '0');

/**
 * @param {Date} date
 * @returns {string}
 */
const formatDate = (date) => {
  return `${date.getFullYear()}-${toTwoDigits(date.getMonth() + 1)}-${toTwoDigits(date.getDate())}`;
};

/**
 * @param {Date} date
 * @returns {string}
 */
const formatDateTime = (date) => {
  return `${formatDate(date)}T${toTwoDigits(date.getHours())}:${toTwoDigits(date.getMinutes())}:${toTwoDigits(date.getSeconds())}`;
}

function importRevenue() {
  // column indices on sheet
  const indexBusinessDate = 0;
  const indexTime = 1;
  const indexNet = 2;
  const indexTax = 3;
  const indexGross = 4;
  const indexWorkSpace = 5;

  const file = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const setupSheet = file.getSheetByName('Setup');
  if (!setupSheet) throw new Error('Sheet named "Setup" not found');
  const apiSheet = file.getSheetByName('API');
  if (!apiSheet) throw new Error('Sheet named "API" not found');
  const revenueSheet = file.getSheetByName('Hourly Revenue');
  if (!revenueSheet) throw new Error('Sheet named "Hourly Revenue" not found');

  const setupData = setupSheet.getDataRange().getValues();
  const apiData = apiSheet.getDataRange().getValues();
  const revenueData = revenueSheet.getDataRange().getValues();

  const apiKey = setupData.find(row => row[0] === 'Api Key')[1];
  const apiUrl = apiData.find(row => row[0] === 'Api URL')[1];

  // if (businessDate instanceof Date) {
  //   businessDate = formatDate(businessDate);
  // }

  /**
   * @typedef {object} Amount
   * @prop {number} net
   * @prop {number} tax
   * @prop {number} gross
   *
   * @typedef {object} RevenueEntry
   * @prop {string} time
   * @prop {Amount} amount
   */

  const /** @type {Record<string, RevenueEntry[]>} */ entriesByDateAndWorkSpace = {};
  const /** @type {Record<string, Amount>} */ totalsByDateAndWorkSpace = {};

  for (let i = 1; i < revenueData.length; i++) {
    const row = revenueData[i];

    const businessDate = row[indexBusinessDate] instanceof Date ? formatDate(row[indexBusinessDate]) : row [indexBusinessDate];
    const byDateSpace = `${businessDate},${row[indexWorkSpace]}`;

    let entriesForGroup = entriesByDateAndWorkSpace[byDateSpace];
    if (!entriesForGroup) {
      entriesForGroup = [];
      entriesByDateAndWorkSpace[byDateSpace] = entriesForGroup;
    }
    let totalsForGroup = totalsByDateAndWorkSpace[byDateSpace];
    if (!totalsForGroup) {
      totalsForGroup = { net: 0, tax: 0, gross: 0 };
      totalsByDateAndWorkSpace[byDateSpace] = totalsForGroup;
    }

    const entry = {
      time: row[indexTime] instanceof Date ? formatDateTime(row[indexTime]) : row[indexTime],
      amount: { net: +row[indexNet], tax: +row[indexTax], gross: +row[indexGross] },
    };
    entriesForGroup.push(entry);
    totalsForGroup.net += entry.amount.net;
    totalsForGroup.tax += entry.amount.tax;
    totalsForGroup.gross += entry.amount.gross;
  }

  const ok = [];
  const errors = [];

  for (const [byDateSpace, entriesForGroup] of Object.entries(entriesByDateAndWorkSpace)) {
    const [businessDate, workspace] = byDateSpace.split(',');

    const postBody = {
      businessDate,
      workspace: { externalNumber: +workspace },
      total: totalsByDateAndWorkSpace[byDateSpace],
      hourly: entriesForGroup,
    };

    // ui.alert(JSON.stringify(postBody, undefined, 2));
    // continue;

    const postResult = PostJSONApiKey(`${apiUrl}revenue/hourly`, apiKey, postBody);

    if (postResult === 'ok') {
      ok.push([businessDate, workspace]);
      // ui.alert('upload succesfull');
    } else {
      errors.push([businessDate, workspace, postResult]);
      // ui.alert(`There were errors when uploading the file:\n${postResult}`);
    }
  }

  let message = '';
  if (ok.length) {
    message += `Sucessfully uploaded the following revenue data:\n- ${ok.map(([date, space]) => `workspace ${space} on ${date}`).join('\n- ')}\n\n`;
  }
  if (errors.length) {
    message += `There where errors:\n- ${errors.map(([date, space, err]) => `on ${date} for workspace ${space}:\n  ${err}`).join('\n- ')}`;
  }
  ui.alert(message);
}

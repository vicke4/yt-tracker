/* eslint-disable no-use-before-define */

/**
 * Returns alignment array
 *
 * @param {Number} noOfRows the alignment to be applied.
 * @param {Number} noOfColumns the alignment to be applied.
 * @param {String} format the format value to be returned.
 * @return {Array} for vertical or horizontal alignment.
 */
export const getCellFormat = (noOfRows, noOfColumns, format) => {
  const values = [];

  for (let i = 0; i < noOfRows; i += 1) {
    const value = [];

    for (let j = 0; j < noOfColumns; j += 1) {
      value.push(format);
    }

    values.push(value);
  }

  return values;
};

export const getSheetInfo = () => {
  const sheetInfo = getSetProperty('sheetInfo', 'user', 'json');

  if (sheetInfo) {
    sheetInfo.disableThumbnail = getSetProperty('disableThumbnail', 'user', 'bool');
    sheetInfo.abbreviate = getSetProperty('abbreviate', 'user', 'bool');
  }

  return sheetInfo;
};

/**
 * Function that gets the property of the document.
 *
 * @param {String} key to be used get the property of the document.
 * @param {String} type to be used get the type of property.
 * valid types - user, script, document
 * @param {String} data type of the property, possible values: json, bool, string.
 * @param {String} action to be performed, get/set of the property.
 * @param {String} value if action === 'set', value to be set.
 * @return {String} if action === 'get' else null;
 */

export const getSetProperty = (key, type, dataType, action, value) => {
  let properties;
  let val;

  /* Added to handle Exception: Argument too large: value error */
  if (key === 'channelUploadMapping') {
    if (action === 'set') {
      const keys = Object.keys(value);
      if (keys.length > 50) return setChannelUploadMapping(value, keys);
    }

    const channelCount = getSetProperty('channelListCount', 'user');

    if (channelCount) {
      return recursiveFetchChannelUploadMapping();
    }
  }
  /* error handling finishes */

  if (type === 'user') {
    properties = PropertiesService.getUserProperties();
  } else if (type === 'script') {
    properties = PropertiesService.getScriptProperties();
  } else {
    properties = PropertiesService.getDocumentProperties();
  }

  if (action === 'set') {
    properties.setProperty(key, dataType === 'json' ? JSON.stringify(value) : value);
  } else {
    const propertyValue = properties.getProperty(key);
    if (!propertyValue) return null;

    if (dataType === 'json') val = JSON.parse(propertyValue);
    else if (dataType === 'bool') val = propertyValue === 'true';
    else val = propertyValue;
  }

  return val;
};

const insertHeader = (headerValues, sheet, columnCount, type) => {
  //  Logger.log(headerValues);
  const sheetRange = sheet.getRange(type === 1 ? 1 : 2, 1, 1, columnCount);
  const hexColour = type === 1 ? '#FF0000' : '#00FF00';

  if (type === 1) {
    sheetRange
      .mergeAcross()
      .setFontSizes(getCellFormat(1, columnCount, '11'))
      .setFontColor('white');
  }

  sheetRange
    .setHorizontalAlignments(getCellFormat(1, columnCount, 'center'))
    .setFontFamilies(getCellFormat(1, columnCount, 'Merriweather'))
    .setFontWeight('bold')
    .setVerticalAlignments(getCellFormat(1, columnCount, 'middle'))
    .setBackgrounds(getCellFormat(1, columnCount, hexColour))
    .setValues(headerValues);
};

const setColumnWidths = (sheet, headerValues, type) => {
  sheet.setColumnWidth(1, type === 'Channels' ? 200 : 110);
  sheet.setColumnWidth(2, type === 'Channels' ? 150 : 425);

  let countColumnStart = 4;

  if (headerValues.indexOf('Thumbnail') > -1) sheet.setColumnWidth(3, 130);
  else countColumnStart = 3;

  sheet.setColumnWidths(countColumnStart, headerValues.length - countColumnStart + 1, 115);
};

export const insertHeaders = (sheet, type, sheetInf, titl) => {
  let title = titl;
  let sheetInfo = sheetInf;
  if (!title) title = `YouTube ${type}`;

  if (!sheetInfo) sheetInfo = { disableThumbnail: false, abbreviate: false };
  else reset(sheet);

  const maxColumns = sheet.getMaxColumns();
  if (maxColumns > 9) {
    sheet.deleteColumns(9, maxColumns - 9);
  }

  const headerTwo = [[`${type === 'Channels' ? 'Channel ' : 'Video '}ID`, 'Title']];

  if (!sheetInfo.disableThumbnail) headerTwo[0].push('Thumbnail');

  headerTwo[0].push('Published on');

  if (type === 'Channels') headerTwo[0].push('# of videos');
  else headerTwo[0].push('Duration');

  headerTwo[0].push('# of views');

  if (type === 'Channels') headerTwo[0].push('# of subscribers');
  else headerTwo[0].push('# of likes', '# of dislikes', '# of comments');

  sheet.setRowHeight(1, 50);

  const columnCount = headerTwo[0].length;
  const headerOne = getCellFormat(1, columnCount, ' ');
  headerOne[0][0] = title;

  setColumnWidths(sheet, headerTwo[0], type);
  insertHeader(headerOne, sheet, columnCount, 1);
  insertHeader(headerTwo, sheet, columnCount, 2);
  sheet.setFrozenRows(2);
};

export const setChannelUploadMapping = (value, keys) => {
  let objChunk = {};
  let index = 0;

  for (let i = 0; i < keys.length; i += 1) {
    objChunk[keys[i]] = value[keys[i]];

    if (i % 50 === 0 && i !== 0) {
      getSetProperty(`channelUploadMapping${index}`, 'user', 'json', 'set', objChunk);
      index += 1;
      objChunk = {};
    } else if (i === keys.length - 1) {
      getSetProperty(`channelUploadMapping${index}`, 'user', 'json', 'set', objChunk);
      index += 1;
    }
  }

  if (index > 0) getSetProperty('channelListCount', 'user', null, 'set', index);
};

export const extend = (obj, src) => {
  // for (const key in src) {
  //   if (src.hasOwnProperty(key)) obj[key] = src[key];
  // }

  Object.keys(src).forEach(key => {
    obj[key] = src[key];
  });

  return obj;
};

export const recursiveFetchChannelUploadMapping = () => {
  const channelCount = +getSetProperty('channelListCount', 'user');
  let returnObject = {};
  let chunkJSON;

  for (let i = 0; i < channelCount; i += 1) {
    chunkJSON = getSetProperty(`channelUploadMapping${i}`, 'user', 'json');
    returnObject = extend(returnObject, chunkJSON);
  }

  return returnObject;
};

/**
 * Function that toggles the given properties key.
 *
 * @param {String} key properies key to be set or unset.
 * @param {String} type to be used get the type of property.
 * valid types - user, script, document
 * @return {Boolean}
 */
export const toggle = (key, type) => {
  let properties = null;

  if (type === 'user') {
    properties = PropertiesService.getUserProperties();
  } else if (type === 'script') {
    properties = PropertiesService.getScriptProperties();
  } else {
    properties = PropertiesService.getDocumentProperties();
  }

  const value = properties.getProperty(key) !== 'true';
  properties.setProperty(key, value);
  return value;
};

/**
 * Returns the Array of sheet ids.
 *
 * @param {SpreadsheetObject} from which the sheet ids to be found.
 * @return {Array}
 */
export const getSheetIds = spreadsheet => {
  const sheets = spreadsheet.getSheets();
  const sheetIds = sheets.map(sheet => sheet.getSheetId());

  return sheetIds;
};

/**
 * Returns the Object of with sheet.
 *
 * @param {SpreadsheetObject} from which the sheet ids to be found.
 * @param {sheetId} ID of the sheet needed.
 * @return {Object}
 */
export const getSheet = (ss, sheetId, spreadsheetId) => {
  let sheet;
  let sheets;
  let spreadsheet = ss;

  try {
    if (!spreadsheet) {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    }

    sheets = spreadsheet.getSheets();
  } catch (e) {
    sheets = [];
  }

  for (let i = 0; i < sheets.length; i += 1) {
    if (sheets[i].getSheetId() === sheetId) {
      sheet = sheets[i];
      break;
    }
  }

  return sheet;
};

/**
 * Returns the abbreviated string
 *
 * @param {count}, integer to be converted Eg: 1000000
 * @param {withAbbr}, whether the symbol must be appended with string
 * @param {decimals}, digits needed after decimal point
 * @return {string} Eg: 1M
 */
export const formatCount = (count, isWithAbbr, decimalPoint) => {
  const COUNT_ABBRS = ['', 'K', 'M', 'B', 'Q', 'P', 'E', 'Z', 'Y'];
  const { log, floor } = Math;
  const withAbbr = isWithAbbr || true;
  const decimals = decimalPoint || 2;

  const i = count === 0 ? count : floor(log(count) / log(1000));
  let result = parseFloat((count / 1000 ** i).toFixed(decimals));
  if (withAbbr) {
    result += COUNT_ABBRS[i];
  }
  return result;
};

/**
 * Returns the readable number format
 *
 * @param {interger} Eg: 1000000
 * @return {string} Eg: 1,000,000
 */
export const formatNumber = (number, abbreviate) => {
  let readableInt;

  if (abbreviate) readableInt = formatCount(number);
  else readableInt = String(number).replace(/\B(?=(\d{3})+\b)/g, ',');

  return readableInt;
};

/**
 * Resets the sheet
 *
 */
export const reset = sheet => {
  sheet.clear({ formatOnly: true, contentsOnly: true });
};

/**
 * Returns the HTML for open tracking spreadsheet.
 *
 * @return {HTMLOutput}
 */
export const getTrackSheetLinkHTML = sheetObj => {
  const t = HtmlService.createTemplateFromFile('openAnotherSheet');
  t.url = sheetObj.url;
  return t
    .evaluate()
    .setHeight(100)
    .setWidth(270);
};

export const getChannelVideoCountHTML = count => {
  const t = HtmlService.createTemplateFromFile('channelVideoLimit');
  t.count = count;
  return t
    .evaluate()
    .setHeight(120)
    .setWidth(270);
};

/* eslint-disable no-unused-vars */

import {
  getCellFormat,
  getChannelVideoCountHTML,
  getSetProperty,
  getSheet,
  getSheetInfo,
  getTrackSheetLinkHTML,
  insertHeaders,
  reset,
  toggle
} from '../utils';
import { addChannel, addVideo, fetchChannelVideos } from './channelVideo';

/**
 * function that sets the addon menu=> .
 *
 * @param {Object} e event object with authentication info.
 */
export const setMenuItems = e => {
  const menu = SpreadsheetApp.getUi().createAddonMenu();
  const menuObj = [
    { name: 'Track YouTube Videos', functionName: 'trackVideos' },
    { name: 'Track YouTube Channels', functionName: 'trackChannels' }
  ];

  if (e && e.authMode !== ScriptApp.AuthMode.NONE) {
    const sheetInfo = getSetProperty('sheetInfo', 'user', 'json');
    if (!sheetInfo) getSetProperty('autoRefresh', 'user', 'bool', 'set', true);
    const addExtraMenu =
      sheetInfo && sheetInfo.spreadsheetId === SpreadsheetApp.getActive().getId();
    const abbreviate = getSetProperty('abbreviate', 'user', 'bool');
    const autoRefresh = getSetProperty('autoRefresh', 'user', 'bool');
    const disableThumbnail = getSetProperty('disableThumbnail', 'user', 'bool');

    if (addExtraMenu) {
      menuObj.push(
        { name: 'Fetch Channel Videos', functionName: 'fetchChannelVideos' },
        { name: 'Change Channel fetch Video count', functionName: 'changeChannelVideoCount' },
        null,
        { name: 'Abbreviate counts', functionName: 'toggleAbbreviation', propertyKey: abbreviate },
        {
          name: 'Auto Refresh on open',
          functionName: 'toggleAutoRefresh',
          propertyKey: autoRefresh
        },
        {
          name: 'Disable Thumbnail',
          functionName: 'toggleThumbnail',
          propertyKey: disableThumbnail
        },
        null,
        { name: 'Refresh', functionName: 'refresh' }
      );
    }
  }

  menuObj.forEach(mObj => {
    if (!mObj) {
      menu.addSeparator();
      return;
    }

    if (mObj.propertyKey) {
      mObj.name = `✓ ${mObj.name}`;
    }

    menu.addItem(mObj.name, mObj.functionName);
  });

  menu.addToUi();
};

/**
 * Creates triggers.
 *
 */
const setTrigger = type => {
  try {
    if (type === 'edit') {
      ScriptApp.newTrigger('onEditTrigger')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onEdit()
        .create();
    } else if (type === 'open') {
      ScriptApp.newTrigger('onOpenTrigger')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onOpen()
        .create();
    }
  } catch (e) {
    Logger.log('Error while setting trigger', e);
  }
};

const insertNewSheet = (spreadsheet, type, sheetInf) => {
  let sheetInfo = sheetInf;
  const sheetName = `YouTube ${type}`;
  if (!sheetInfo)
    sheetInfo = {
      spreadsheetId: spreadsheet.getId(),
      videos: null,
      channels: null
    };
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  const sheetId = sheet.getSheetId();
  sheetInfo[type.toLowerCase()] = {
    sheetId,
    url: `${spreadsheet.getUrl()}#gid=${sheetId}`
  };

  insertHeaders(sheet, type, sheetInfo, `YouTube ${type}`);
  getSetProperty('sheetInfo', 'user', 'json', 'set', sheetInfo);
  return sheet;
};

const activateSheet_ = type => {
  const sheetInfo = getSheetInfo();
  const spreadsheet = SpreadsheetApp.getActive();
  const isChannel = type === 'Channels';
  const getSheetDisplayToast = () => {
    spreadsheet.toast(
      `
        Add ${isChannel ? 'channel ids' : 'video ids'} to first column of this sheet to start
        tracking. Example: ${isChannel ? 'UCwppdrjsBPAZg5_cUwQjfMQ' : 'kJQP7kiw5Fk'}
      `,
      '',
      -1
    );

    return insertNewSheet(spreadsheet, type, sheetInfo);
  };
  const checkSetTriggers = () => {
    const autoRefresh = getSetProperty('autoRefresh', 'user', 'bool');
    const triggerList = ScriptApp.getUserTriggers(spreadsheet).map(t => t.getHandlerFunction());
    const openTrigger = triggerList.indexOf('onOpenTrigger') === -1 && autoRefresh;
    const editTrigger = triggerList.indexOf('onEditTrigger') === -1;

    if (openTrigger) setTrigger('open');
    if (editTrigger) setTrigger('edit');
    if (openTrigger || editTrigger) setMenuItems({ authMode: ScriptApp.AuthMode.FULL });
  };

  let sheet;
  let newTrackingSheet;

  if (sheetInfo && sheetInfo.spreadsheetId && sheetInfo.spreadsheetId !== spreadsheet.getId()) {
    let actualTrackerType = sheetInfo[type.toLowerCase()] ? type.toLowerCase() : 'shuffle';
    if (actualTrackerType === 'shuffle') {
      actualTrackerType = isChannel ? 'videos' : 'channels';
    }

    newTrackingSheet = true;
    const trackerSheet = getSheet(
      null,
      sheetInfo[actualTrackerType].sheetId,
      sheetInfo.spreadsheetId
    );

    if (trackerSheet) {
      SpreadsheetApp.getUi().showModalDialog(
        getTrackSheetLinkHTML(sheetInfo[actualTrackerType.toLowerCase()]),
        ' '
      );
      setMenuItems({ authMode: ScriptApp.AuthMode.FULL });
      return true;
    }
  }

  if (!sheetInfo || !sheetInfo[type.toLowerCase()]) {
    sheet = getSheetDisplayToast();
    checkSetTriggers();
  } else {
    sheet = getSheet(spreadsheet, sheetInfo[type.toLowerCase()].sheetId);

    if (!sheet) {
      if (newTrackingSheet) sheetInfo.spreadsheetId = spreadsheet.getId();
      checkSetTriggers();

      sheet = getSheetDisplayToast();
    }
  }

  if (newTrackingSheet) setMenuItems({ authMode: ScriptApp.AuthMode.FULL });
  return sheet.activate();
};

export const trackVideos = () => {
  activateSheet_('Videos');
};

export const trackChannels = () => {
  activateSheet_('Channels');
};

export const changeChannelVideoCount = () => {
  activateSheet_('Channels');
  const count = getSetProperty('channelVideoCount', 'user') || 500;

  try {
    SpreadsheetApp.getUi().showModalDialog(
      getChannelVideoCountHTML(count),
      'Number of channel videos to fetch'
    );
  } catch (e) {
    Logger.log(e);
  }
};

const getValidIds = (sheet, fromChannelVideos) => {
  const rowCount = sheet.getLastRow();
  const range = sheet.getRange(3, 1, rowCount, 1);
  const idRangeValues = range.getValues();
  const idList = [];

  for (let i = 0; i < idRangeValues.length; i += 1) {
    if (idRangeValues[i][0] !== '') idList.push(idRangeValues[i][0]);
  }

  if (!fromChannelVideos) range.setValues(getCellFormat(rowCount, 1, ''));

  return idList;
};

const refreshChannelVideoSheet = (spreadsheet, sheet, type, sheetInfo, replaceHeaders, idLst) => {
  const idList = idLst || getValidIds(sheet, true);
  const channelUploadMapping = getSetProperty('channelUploadMapping', 'user', 'json');

  for (let i = 0; i < idList.length; i += 1) {
    const channelInfo = channelUploadMapping[idList[i]];
    const requested = channelInfo ? channelInfo.requested : null;

    if (requested) {
      const channelSheet = getSheet(spreadsheet, channelInfo.sheetId);

      if (channelSheet && replaceHeaders) {
        reset(channelSheet);
        insertHeaders(channelSheet, type, sheetInfo, channelInfo.title);
      }

      fetchChannelVideos(idList[i], channelSheet, true);
    }
  }
};

const refreshSheet = (spreadsheet, type, sheetInfo, replaceHeaders, idList, fromTrigger) => {
  const sheeType = type === 'ChannelVideos' ? 'Channels' : type;
  const sheetObj = sheetInfo[sheeType.toLowerCase()];
  const sheet = sheetObj ? getSheet(spreadsheet, sheetObj.sheetId) : null;

  if (sheet && type === 'ChannelVideos') {
    refreshChannelVideoSheet(spreadsheet, sheet, type, sheetInfo, replaceHeaders, idList);
    return;
  }

  if (sheet) {
    //    var ids = getValidIds(sheet);
    let ids = getValidIds(sheet).reduce((accum, currentValue) => {
      if (accum === '') return accum + currentValue;
      return `${accum},${currentValue}`;
    }, '');

    if (replaceHeaders) {
      reset(sheet);
      insertHeaders(sheet, type, sheetInfo);
    }

    const func = type === 'Videos' ? addVideo : addChannel;
    //    if (ids.length > 0) func(sheet, ids.join(), sheetInfo, 3, true);
    if (ids.length > 0) {
      let i;
      let j;
      let temparray;

      const chunk = 50;
      const array = ids.split(',');
      let cellPosition = 3;
      for (i = 0, j = array.length; i < j; i += chunk) {
        temparray = array.slice(i, i + chunk);
        ids = temparray.join();
        func(sheet, ids, sheetInfo, cellPosition, true);

        cellPosition += chunk;
      }
    } else if (!fromTrigger)
      spreadsheet.toast(`Please add some ${type.toLowerCase()} before trying to refresh`, '', 10);
  }
};

const refreshSheets = (spreadsheet, sheetInfo, replaceHeaders, fromTrigger) => {
  ['Videos', 'Channels', 'ChannelVideos'].forEach(type => {
    refreshSheet(spreadsheet, type, sheetInfo, replaceHeaders, null, fromTrigger);
  });

  setMenuItems({ authMode: ScriptApp.AuthMode.FULL });
};

/**
 * function that toggles autorefresh functionality of the addo=> n
 * on sheet open.
 */
export const toggleAutoRefresh = () => {
  const sheetInfo = getSheetInfo();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (sheetInfo.spreadsheetId !== spreadsheet.getId()) {
    setMenuItems({ authMode: ScriptApp.AuthMode.FULL });
    SpreadsheetApp.getUi().showModalDialog(getTrackSheetLinkHTML(sheetInfo.videos.url), ' ');
    return;
  }

  if (toggle('autoRefresh', 'user')) {
    ScriptApp.newTrigger('onOpenTrigger')
      .forSpreadsheet(spreadsheet.getId())
      .onOpen()
      .create();
  } else {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getEventType() === ScriptApp.EventType.ON_OPEN) ScriptApp.deleteTrigger(trigger);
    });
  }

  setMenuItems({ authMode: ScriptApp.AuthMode.FULL });
};

export const toggleThumbnail = () => {
  const sheetInfo = getSheetInfo();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (sheetInfo.spreadsheetId !== spreadsheet.getId()) {
    setMenuItems({ authMode: ScriptApp.AuthMode.FULL });
    SpreadsheetApp.getUi().showModalDialog(getTrackSheetLinkHTML(sheetInfo.videos.url), ' ');
    return;
  }

  sheetInfo.disableThumbnail = toggle('disableThumbnail', 'user');
  refreshSheets(spreadsheet, sheetInfo, true, true);
};

export const toggleAbbreviation = () => {
  const sheetInfo = getSheetInfo();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (sheetInfo.spreadsheetId !== spreadsheet.getId()) {
    setMenuItems({ authMode: ScriptApp.AuthMode.FULL });
    SpreadsheetApp.getUi().showModalDialog(getTrackSheetLinkHTML(sheetInfo.videos.url), ' ');
    return;
  }

  sheetInfo.abbreviate = toggle('abbreviate', 'user');
  refreshSheets(spreadsheet, sheetInfo, true, true);
};

export const refresh = () => {
  const sheetInfo = getSheetInfo();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (sheetInfo.spreadsheetId !== spreadsheet.getId()) {
    setMenuItems({ authMode: ScriptApp.AuthMode.FULL });
    SpreadsheetApp.getUi().showModalDialog(getTrackSheetLinkHTML(sheetInfo.videos.url), ' ');
    return;
  }

  const activeSheetId = spreadsheet.getActiveSheet().getSheetId();
  const videoSheetId = sheetInfo.videos ? sheetInfo.videos.sheetId : null;
  if (videoSheetId === activeSheetId) {
    refreshSheet(spreadsheet, 'Videos', sheetInfo);
    return;
  }

  const channelSheetId = sheetInfo.channels ? sheetInfo.channels.sheetId : null;
  if (channelSheetId === activeSheetId) {
    refreshSheet(spreadsheet, 'Channels', sheetInfo);
    return;
  }

  const channelUploadMapping = getSetProperty('channelUploadMapping', 'user', 'json');
  if (channelUploadMapping) {
    const channelIds = Object.keys(channelUploadMapping);
    for (let i = 0; i < channelIds.length; i += 1) {
      const channelId = channelIds[i];

      if (channelUploadMapping[channelId].sheetId === activeSheetId) {
        refreshSheet(spreadsheet, 'ChannelVideos', sheetInfo, false, [channelId]);
      }
    }
  }
};

/**
 * Simple trigger that runs on sheet open.
 *
 */
export const onOpenTrigger = () => {
  const cache = CacheService.getUserCache();

  // Returns if called by trigger and the 6h cache is not expired
  if (cache.get('autoRefreshCacheKey')) {
    return;
  }

  // 6h cache to not refresh links sheet frequently on open
  cache.put('autoRefreshCacheKey', true, 21600);

  const sheetInfo = getSheetInfo();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  refreshSheets(spreadsheet, sheetInfo, null, true);
};

/**
 * Simple trigger that runs on sheet edit.
 *
 */
export const onEditTrigger = e => {
  try {
    const sheetInfo = getSheetInfo();
    const sheetIdArr = [
      sheetInfo.videos ? sheetInfo.videos.sheetId : null,
      sheetInfo.channels ? sheetInfo.channels.sheetId : null
    ];
    const sheetIndex = sheetIdArr.indexOf(e.source.getActiveSheet().getSheetId());
    const flag = null;

    if (sheetIndex === -1) return;

    const cell = e.range.getA1Notation();

    if (cell[0] !== 'A') return;
    if (cell.length === 2 && e.value === '') return;

    const startingRow = e.range.getRow();
    const lastRow = e.range.getLastRow() - startingRow + 1;

    if (startingRow < 3) {
      return;
    }
    const incomingChanges = e.range.getValues();

    for (let i = 0; i < incomingChanges.length; i += 1) {
      if (incomingChanges[i][0] === '') {
        return;
      }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('Fetching data please wait...', '', 1);

    const sheet = SpreadsheetApp.getActiveSheet();
    const activeRange = sheet.getRange(startingRow, 1, lastRow, 1);
    const values = activeRange.getValues();
    const emptyValues = [];
    let ids = values.reduce((accum, currentValue) => {
      emptyValues.push(['']);
      if (accum === '') return accum + currentValue[0];
      return `${accum},${currentValue[0]}`;
    }, '');

    activeRange.setValues(emptyValues);

    let i;
    let j;
    let temparray;
    const chunk = 50;
    const array = ids.split(',');
    let cellPosition = Number(cell.split(':')[0].substr(1));

    for (i = 0, j = array.length; i < j; i += chunk) {
      temparray = array.slice(i, i + chunk);
      ids = temparray.join();

      if (sheetIndex === 0) addVideo(sheet, ids, sheetInfo, cellPosition);
      else addChannel(sheet, ids, sheetInfo, cellPosition);

      cellPosition += chunk;
    }
  } catch (er) {
    Logger.log(er);
  }
};

/**
 * Simple trigger that runs on sheet open.
 *
 */
export const onOpen = e => {
  try {
    setMenuItems(e);
  } catch (error) {
    //    Logger.log(e);
  }
};

/**
 * Simple trigger that runs upon addon installation.
 *
 */
export const onInstall = e => {
  onOpen(e);
};

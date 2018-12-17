import {
  formatNumber,
  getCellFormat,
  getSetProperty,
  getSheet,
  getSheetInfo,
  insertHeaders
} from '../utils';

const updateSheet = (sheet, sheetValues, disableThumbnail, insertRow) => {
  try {
    const columnCount = sheetValues[0].length;
    const rowCount = sheetValues.length;

    sheet
      .getRange(insertRow, 1, rowCount, columnCount)
      .setHorizontalAlignments(getCellFormat(rowCount, columnCount, 'center'))
      .setFontFamilies(getCellFormat(rowCount, columnCount, 'Roboto'))
      .setVerticalAlignments(getCellFormat(rowCount, columnCount, 'middle'))
      .setValues(sheetValues);

    sheet.setRowHeights(insertRow, rowCount, disableThumbnail ? 30 : 90);

    sheet
      .getRange(insertRow, 1, rowCount, 2)
      .setHorizontalAlignments(getCellFormat(rowCount, 2, 'left'));
  } catch (e) {
    Logger.log('Error');
    Logger.log(e);
  }
};

const getResponse = (type, ids) => {
  let fields;
  let api;

  if (type === 'video') {
    fields =
      `items(id,contentDetails/duration,snippet(publishedAt,` +
      `thumbnails(default,medium,standard),title),statistics),nextPageToken`;
    api = YouTube.Videos;
  } else {
    fields =
      `items(contentDetails/relatedPlaylists/uploads,` +
      `id,snippet(publishedAt,thumbnails/medium,title),` +
      `statistics(subscriberCount,videoCount,viewCount))`;
    api = YouTube.Channels;
  }

  return api.list('contentDetails,snippet,statistics', {
    id: ids,
    fields
  });
};

// const formatDuration = duration => {
//   const durArr = duration
//     .replace('PT', '')
//     .replace('H', ':')
//     .replace('S', '')
//     .replace('M', ':')
//     .split(':');
//   let tempTxt;

//   if (durArr.length > 1) {
//     for (let i = durArr.length - 1; i > 0; i--) {
//       tempTxt = durArr[i].length === 1 ? 0 + durArr[i] : durArr[i];
//       durArr[i] = tempTxt === '' ? '00' : tempTxt;
//     }
//   }

//   return durArr.join(':');
// };

const formatDurationNew = duration => {
  const dur = duration.replace('PT', '');
  const hArr = dur.split('H');
  let mArr = dur.split('M');
  let sArr = dur.split('S');
  let sTemp;
  let mTemp;

  if (sArr.length > 1) {
    sArr = sArr[0].split('H');
    sArr = sArr[sArr.length - 1].split('M');
    sTemp = Number(sArr[sArr.length - 1]);
  } else sTemp = '';

  sTemp = sTemp !== '' ? String(sTemp - 1) : 59;
  sTemp = typeof sTemp === 'string' && sTemp.length === 1 ? 0 + sTemp : sTemp;

  if (mArr.length > 1) {
    mArr = mArr[0].split('H');
    mTemp = Number(mArr[mArr.length - 1]);
  } else mTemp = '';

  let hTemp = hArr.length > 1 ? Number(hArr[0]) : '';

  if (sTemp === 59 && mTemp === '') {
    if (hTemp === '') mTemp = '00';
    else mTemp = 59;
  } else if (sTemp === 59 && mTemp !== '') {
    mTemp = String(mTemp - 1);
  } else {
    mTemp = mTemp === '' ? '00' : String(mTemp);
  }

  mTemp = typeof mTemp === 'string' && mTemp.length === 1 ? 0 + mTemp : mTemp;

  if (sTemp === 59 && mTemp === 59 && hTemp >= 1) hTemp -= 1;
  if (hTemp === 0 || hTemp === '') return `${mTemp} : ${sTemp}`;

  return `${hTemp} : ${mTemp} : ${sTemp}`;
};

const updateSheetValuesVidoes = (items, sheetValues, abbreviate, disableThumbnail) => {
  if (items && items.length > 0) {
    let sheetIterVar = sheetValues.length;

    for (let i = 0; i < items.length; i += 1) {
      sheetValues.push([
        `=hyperlink("https://www.youtube.com/watch?v=${items[i].id}","${items[i].id}")`,
        items[i].snippet.title
      ]);

      if (!disableThumbnail) {
        sheetValues[sheetIterVar].push(
          `=IMAGE("${items[i].snippet.thumbnails.medium.url}", 4, 73, 130)`
        );
      }

      sheetValues[sheetIterVar].push(items[i].snippet.publishedAt.split('T')[0]);

      sheetValues[sheetIterVar].push(
        formatDurationNew(items[i].contentDetails.duration),
        formatNumber(items[i].statistics.viewCount, abbreviate),
        formatNumber(items[i].statistics.likeCount, abbreviate),
        formatNumber(items[i].statistics.dislikeCount, abbreviate),
        formatNumber(items[i].statistics.commentCount, abbreviate)
      );

      sheetIterVar += 1;
    }
  }
};

export const addVideo = (sheet, ids, sheetInfo, insertRow) => {
  const sheetValues = [];
  let i;

  let j;

  const chunk = 50;

  const idList = ids.split(',');

  for (i = 0, j = idList.length; i < j; i += chunk) {
    const videoIds = idList.slice(i, i + chunk).join();
    const response = getResponse('video', videoIds);

    updateSheetValuesVidoes(
      response.items,
      sheetValues,
      sheetInfo.abbreviate,
      sheetInfo.disableThumbnail
    );
  }

  if (sheetValues.length > 0)
    updateSheet(sheet, sheetValues, sheetInfo.disableThumbnail, insertRow);
};

export const addChannel = (sheet, ids, sheetInfo, insertRow, fromRefresh) => {
  const response = getResponse('channel', ids);
  const sheetValues = [];
  const { items: respItems } = response;
  const { abbreviate } = sheetInfo;
  const channelUploadMapping = getSetProperty('channelUploadMapping', 'user', 'json') || {};

  // Ordering response according to input
  let items = [];
  const idList = ids.split(',');

  respItems.forEach(i => {
    items[idList.indexOf(i.id)] = i;
  });

  items = items.filter(item => item);

  if (items && items.length > 0) {
    for (let i = 0; i < items.length; i += 1) {
      sheetValues.push([
        `=hyperlink("https://www.youtube.com/channel/${items[i].id}","${items[i].id}")`,
        items[i].snippet.title
      ]);

      if (!channelUploadMapping[items[i].id]) {
        channelUploadMapping[items[i].id] = {
          playlistId: items[i].contentDetails.relatedPlaylists.uploads,
          title: items[i].snippet.title
        };
      }

      if (!fromRefresh) {
        channelUploadMapping[items[i].id].requested = false;
      }

      if (!sheetInfo.disableThumbnail) {
        sheetValues[i].push(`=IMAGE("${items[i].snippet.thumbnails.medium.url}", 4, 80, 80)`);
      }

      sheetValues[i].push(items[i].snippet.publishedAt.split('T')[0]);

      sheetValues[i].push(
        formatNumber(items[i].statistics.videoCount, abbreviate),
        formatNumber(items[i].statistics.viewCount, abbreviate),
        formatNumber(items[i].statistics.subscriberCount, abbreviate)
      );
    }
  }

  try {
    getSetProperty('channelUploadMapping', 'user', 'json', 'set', channelUploadMapping);
  } catch (e) {
    Logger.log(e);
  }

  if (sheetValues.length > 0)
    updateSheet(sheet, sheetValues, sheetInfo.disableThumbnail, insertRow);
};

export const fetchChannelVideos = (channlId, shet, programmaticUpdate) => {
  let resp;
  let channelId = channlId;
  let sheet = shet;
  let pageToken;

  const channelVideoIds = [];
  const channelVideoCount = parseInt(getSetProperty('channelVideoCount', 'user'), 10) || 50;
  const sheetInfo = getSheetInfo();
  const channelUploadMapping = getSetProperty('channelUploadMapping', 'user', 'json') || {};
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet = sheet || SpreadsheetApp.getActiveSheet();

  if (!channlId) {
    const range = sheet.getActiveRange();
    channelId = sheet.getRange(range.getRow(), 1).getValue();
  }

  const playlistId = channelUploadMapping[channelId]
    ? channelUploadMapping[channelId].playlistId
    : null;

  if (
    !sheetInfo.channels ||
    (sheetInfo.channels.sheetId !== sheet.getSheetId() && !programmaticUpdate) ||
    !playlistId
  ) {
    spreadsheet.toast(
      `
        Kindly highlight the channel to fetch videos on YouTube channels sheet
        (Add-ons → YT Tracker → Track YouTube Channels) and use this feature
      `,
      '',
      20
    );
    return;
  }
  const { sheetId } = channelUploadMapping[channelId];
  sheet = sheetId ? getSheet(SpreadsheetApp.getActiveSpreadsheet(), sheetId) : null;

  if (!sheet) {
    const sheetName = `${channelUploadMapping[channelId].title} - Videos`;
    sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) sheet = SpreadsheetApp.getActive().insertSheet(sheetName);
    insertHeaders(sheet, 'Videos', sheetInfo, sheetName);

    channelUploadMapping[channelId].sheetId = sheet.getSheetId();
    channelUploadMapping[channelId].requested = true;
    getSetProperty('channelUploadMapping', 'user', 'json', 'set', channelUploadMapping);
  }

  if (!programmaticUpdate) sheet.activate();

  do {
    resp = YouTube.PlaylistItems.list('snippet,contentDetails', {
      maxResults: 50,
      playlistId,
      pageToken,
      fields: 'items/snippet/resourceId/videoId,nextPageToken'
    });

    const { items } = resp;

    for (let i = 0; i < items.length; i += 1) {
      channelVideoIds.push(items[i].snippet.resourceId.videoId);
    }

    if (channelVideoIds.length > channelVideoCount) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `
          This channel seems to have more than ${channelVideoCount} videos.
          To fetch more videos, change the limit from add-on menu and try again
        `,
        '',
        -1
      );
      break;
    }

    pageToken = resp.nextPageToken;
  } while (pageToken);

  addVideo(sheet, channelVideoIds.slice(0, channelVideoCount).join(), sheetInfo, 3);
};

export const updateVideoCount = formObject => {
  getSetProperty('channelVideoCount', 'user', null, 'set', formObject.videoCount);
};

import {
  changeChannelVideoCount,
  onEditTrigger,
  onInstall,
  onOpen,
  onOpenTrigger,
  refresh,
  setMenuItems,
  toggleAbbreviation,
  toggleAutoRefresh,
  toggleThumbnail,
  trackChannels,
  trackVideos
} from './core';
import { fetchChannelVideos, updateVideoCount } from './core/channelVideo';
import { getSetProperty } from './utils';

global.changeChannelVideoCount = changeChannelVideoCount;
global.fetchChannelVideos = fetchChannelVideos;
global.onInstall = onInstall;
global.onOpen = onOpen;
global.refresh = refresh;
global.setMenuItems = setMenuItems;
global.toggleAbbreviation = toggleAbbreviation;
global.toggleAutoRefresh = toggleAutoRefresh;
global.toggleThumbnail = toggleThumbnail;
global.trackChannels = trackChannels;
global.trackVideos = trackVideos;
global.updateVideoCount = updateVideoCount;

// Triggers
global.onEditTrigger = onEditTrigger;
global.onOpenTrigger = onOpenTrigger;

// Test ground
global.deleteProp = () => PropertiesService.getUserProperties().deleteAllProperties();
global.getProp = () => {
  const sheetInfo = getSetProperty('sheetInfo', 'user', 'json');
  Logger.log('Check below');
  Logger.log(sheetInfo);
  const cUM = getSetProperty('channelUploadMapping', 'user', 'json');
  Logger.log(cUM);
};

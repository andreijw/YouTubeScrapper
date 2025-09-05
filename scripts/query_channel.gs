const API_KEY = 'API_TOKEN';
const DEFAULT_SUBS = 500000;
const DEFAULT_CHANNELS = 5;
const RADIX = 10;
const TITLE_COLOR = '#d9ead3'

/**
 * Fetches and outputs YouTube channels matching a subject and subscriber threshold.
 */
function getChannelsBySubject(subject, maxSubs, targetCount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clearContents();
  sheet.appendRow(['Channel Title', 'Channel ID', 'Subscribers', 'Niche']);
  styleRows(1, 1, 4, TITLE_COLOR, true, true);

  let collected = 0;
  let nextPageToken = '';

  while (collected < targetCount) {
    const searchData = fetchSearchResults(subject, nextPageToken);
    if (!searchData?.items?.length) break;

    const searchMap = buildSearchMap(searchData.items);
    const channelIds = Object.keys(searchMap);
    const statsData = fetchChannelStats(channelIds);

    for (const channel of statsData.items) {
      if (collected >= targetCount) break;

      const channelId = channel.id;
      const subs = parseInt(channel.statistics.subscriberCount || '0', RADIX);
      if (isNaN(subs) || subs > maxSubs) continue;

      const title = searchMap[channelId]?.title || 'Unknown Title';
      const niche = extractNiche(channel.topicDetails?.topicCategories);

      sheet.appendRow([title, channelId, subs, niche]);
      collected++;
    }

    nextPageToken = searchData.nextPageToken;
    if (!nextPageToken) break;
  }
}

/**
 * Fetches search results for a given subject.
 */
function fetchSearchResults(subject, pageToken = '') {
  const url = `https://www.googleapis.com/youtube/v3/search?part=snippet&type=channel&q=${encodeURIComponent(subject)}&maxResults=10&pageToken=${pageToken}&key=${API_KEY}`;
  const response = UrlFetchApp.fetch(url);
  return JSON.parse(response.getContentText());
}

/**
 * Builds a map of channelId → metadata from search results.
 */
function buildSearchMap(items) {
  const map = {};
  items.forEach(item => {
    const id = item.snippet.channelId;
    map[id] = {
      title: item.snippet.channelTitle
    };
  });
  return map;
}

/**
 * Fetches statistics and topic details for a list of channel IDs.
 */
function fetchChannelStats(channelIds) {
  const url = `https://www.googleapis.com/youtube/v3/channels?part=statistics,topicDetails&id=${channelIds.join(',')}&key=${API_KEY}`;
  const response = UrlFetchApp.fetch(url);
  return JSON.parse(response.getContentText());
}

/**
 * Extracts a readable niche from topic category URLs.
 */
function extractNiche(topicCategories = []) {
  return topicCategories
    .map(url => url.split('/').pop().replace(/_/g, ' '))
    .join(', ') || 'Unknown';
}

/**
 * Prompt the user for the channel subject, max sub counts, and matches
 */
function promptSubject() {
  const ui = SpreadsheetApp.getUi();

  // Prompt for subject
  const subjectResponse = ui.prompt('Enter a subject to search YouTube channels:');
  if (subjectResponse.getSelectedButton() !== ui.Button.OK) return;
  const subject = subjectResponse.getResponseText();

  // Prompt for max subscriber count
  const subsResponse = ui.prompt('Enter maximum subscriber count. Default max_subs (500000):');
  if (subsResponse.getSelectedButton() !== ui.Button.OK) return;
  const input = subsResponse.getResponseText().trim();

  const maxSubs = /^\d+$/.test(input) ? parseInt(input, RADIX) : DEFAULT_SUBS;

  if (isNaN(maxSubs)) {
    ui.alert('Invalid subscriber count. Please enter a number.');
    return;
  }

  const numberResonse = ui.prompt('Number of channels to search. Default (5):');
  if (numberResonse.getSelectedButton() !== ui.Button.OK) return;
  const numberInput = numberResonse.getResponseText().trim();

  const maxChannels = /^\d+$/.test(numberInput) ? parseInt(numberInput, RADIX) : DEFAULT_CHANNELS;

  if (isNaN(maxChannels)) {
    ui.alert('Invalid channel count. Please enter a number.');
    return;
  }

  getChannelsBySubject(subject, maxSubs, maxChannels);
}
const METRIC_COLUMNS = [
  'Videos (1mo)', 'Videos (6mo)',
  'Avg Views (6mo)', 'Avg Likes (6mo)', 'Avg Comments (6mo)',
  'Total Views (6mo)', 'Total Likes (6mo)', 'Total Comments (6mo)',
  'Error'
];

/**
 * Main entry point to collect public YouTube metrics for each channel.
 */
function collectPublicMetrics() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const oneMonthAgo = new Date(now);
  const sixMonthsAgo = new Date(now);
  oneMonthAgo.setMonth(now.getMonth() - 1);
  sixMonthsAgo.setMonth(now.getMonth() - 6);

  // Write headers
  sheet.getRange(1, 5, 1, METRIC_COLUMNS.length).setValues([METRIC_COLUMNS]);
  styleRows(1, 1, 4 + METRIC_COLUMNS.length, TITLE_COLOR, true, true);

  for (let i = 1; i < data.length; i++) {
    const channelId = data[i][1];
    if (!channelId) continue;

    try {
      const videoList = fetchRecentVideos(channelId, sixMonthsAgo.toISOString());
      const stats = fetchVideoStats(videoList);

      const videos1mo = videoList.filter(v => new Date(v.publishedAt) >= oneMonthAgo).length;
      const videos6mo = videoList.length;

      const views = stats.map(s => s.viewCount);
      const likes = stats.map(s => s.likeCount);
      const comments = stats.map(s => s.commentCount);

      const rowMetrics = [
        videos1mo,
        videos6mo,
        average(views),
        average(likes),
        average(comments),
        sum(views),
        sum(likes),
        sum(comments),
        ''
      ];

      sheet.getRange(i + 1, 5, 1, METRIC_COLUMNS.length).setValues([rowMetrics]);
    } catch (err) {
      sheet.getRange(i + 1, 13).setValue(err.message || 'Error');
    }
  }
}

/**
 * Fetches recent videos for a channel since a given ISO date.
 */
function fetchRecentVideos(channelId, publishedAfterISO) {
  const videos = [];
  let nextPageToken = '';

  do {
    const url = `https://www.googleapis.com/youtube/v3/search?part=snippet&channelId=${channelId}&publishedAfter=${publishedAfterISO}&type=video&maxResults=30&pageToken=${nextPageToken}&key=${API_KEY}`;
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());

    if (!data.items) break;

    videos.push(...data.items.map(item => ({
      id: item.id.videoId,
      publishedAt: item.snippet.publishedAt
    })));

    nextPageToken = data.nextPageToken || '';
  } while (nextPageToken);

  return videos;
}

/**
 * Fetches statistics for a list of video IDs.
 */
function fetchVideoStats(videoList) {
  const stats = [];
  const batchSize = 30;

  for (let i = 0; i < videoList.length; i += batchSize) {
    const batch = videoList.slice(i, i + batchSize);
    const ids = batch.map(v => v.id).join(',');
    const url = `https://www.googleapis.com/youtube/v3/videos?part=statistics&id=${ids}&key=${API_KEY}`;
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());

    if (!data.items) continue;

    data.items.forEach(item => {
      const s = item.statistics;
      stats.push({
        viewCount: parseInt(s.viewCount || '0', RADIX),
        likeCount: parseInt(s.likeCount || '0', RADIX),
        commentCount: parseInt(s.commentCount || '0', RADIX)
      });
    });
  }

  return stats;
}

/**
 * Computes average of a numeric array.
 */
function average(arr) {
  return arr.length ? Math.round(sum(arr) / arr.length) : 0;
}

/**
 * Computes sum of a numeric array.
 */
function sum(arr) {
  return arr.reduce((a, b) => a + b, 0);
}
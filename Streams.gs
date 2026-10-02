// THIS CODE IS NOT CURRENTLY USED IN THE EXECUTION DUE TO BUGS AND IS NO LONGER A PRIORITY

var StreamersSheet = SS.getSheetByName("Streamers");

var TWITCH_LINK_RE = /twitch\.tv\/([a-zA-Z0-9_]{4,25})/i;


var STREAM_RECENCY_HOURS = 6;


var streamsCache = null;
var streamsCacheTime = null;
var STREAMS_CACHE_DURATION = 60000; // 1 minute


function refreshStreamers() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('DISCORD_BOT_TOKEN');
  var channelId = props.getProperty('DISCORD_STREAMS_CHANNEL_ID');
  if (!token || !channelId) {
    console.error('refreshStreamers: DISCORD_BOT_TOKEN or DISCORD_STREAMS_CHANNEL_ID not set in Script Properties.');
    return;
  }
  if (!StreamersSheet) {
    console.error('refreshStreamers: no "Streamers" sheet found. Add one with header row: Discord User ID | Discord Username | Twitch Username | First Seen | Last Seen');
    return;
  }

  var messages = fetchDiscordMessages(channelId, token, 100);
  if (!messages) return; // error already logged by fetchDiscordMessages

  mergeStreamersFromMessages(messages);
}


function backfillStreamerHistory(maxMessages) {
  maxMessages = maxMessages || 2000;
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('DISCORD_BOT_TOKEN');
  var channelId = props.getProperty('DISCORD_STREAMS_CHANNEL_ID');
  if (!token || !channelId) {
    console.error('backfillStreamerHistory: DISCORD_BOT_TOKEN or DISCORD_STREAMS_CHANNEL_ID not set in Script Properties.');
    return;
  }
  if (!StreamersSheet) {
    console.error('backfillStreamerHistory: no "Streamers" sheet found.');
    return;
  }

  var before = null;
  var totalFetched = 0;
  var totalFound = 0;

  while (totalFetched < maxMessages) {
    var batch = fetchDiscordMessages(channelId, token, 100, before);
    if (!batch || batch.length === 0) break;

    totalFound += mergeStreamersFromMessages(batch);
    totalFetched += batch.length;
    before = batch[batch.length - 1].id; // Discord returns newest-first; page backward

    if (batch.length < 100) break; // reached the start of the channel
  }

  console.log('backfillStreamerHistory: scanned ' + totalFetched + ' messages, found/updated ' + totalFound + ' streamer link(s).');
}


function fetchDiscordMessages(channelId, token, limit, before) {
  var url = 'https://discord.com/api/v10/channels/' + channelId + '/messages?limit=' + limit;
  if (before) url += '&before=' + before;

  var maxAttempts = 3;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': 'Bot ' + token,
        'User-Agent': 'DiscordBot (https://top100smm2.com, 1.0)'
      },
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    if (code === 200) {
      try {
        return JSON.parse(resp.getContentText());
      } catch (err) {
        console.error('Discord message fetch: could not parse response: ' + err);
        return null;
      }
    }

   
    var isCloudflareBlock = code === 403 && resp.getContentText().indexOf('40333') !== -1;
    if (isCloudflareBlock && attempt < maxAttempts) {
      console.warn('Discord message fetch: got Cloudflare block (40333), retrying (' + attempt + '/' + maxAttempts + ')...');
      Utilities.sleep(1500 * attempt);
      continue;
    }

    console.error('Discord message fetch failed (' + code + '): ' + resp.getContentText());
    return null;
  }

  return null;
}


function mergeStreamersFromMessages(messages) {
  var found = {}; // twitchUsername (lowercase) -> {discordUsername, discordUserId}

  messages.forEach(function (msg) {
    var content = msg.content || '';
    var m = content.match(TWITCH_LINK_RE);
    if (!m) return;
    var uname = m[1].toLowerCase();
  
    if (['videos', 'directory', 'p', 'settings', 'subscriptions', 'wallet', 'friends'].indexOf(uname) !== -1) return;
    found[uname] = {
      discordUsername: (msg.author && (msg.author.global_name || msg.author.username)) || 'Unknown',
      discordUserId: msg.author ? msg.author.id : ''
    };
  });

  var usernames = Object.keys(found);
  if (usernames.length === 0) return 0;

  var data = StreamersSheet.getDataRange().getValues();
  var existingRow = {}; // twitchUsername (lowercase) -> 1-based sheet row
  for (var i = 1; i < data.length; i++) {
    var tw = (data[i][2] || '').toString().trim().toLowerCase();
    if (tw) existingRow[tw] = i + 1;
  }

  var now = new Date();
  usernames.forEach(function (tw) {
    var info = found[tw];
    if (existingRow[tw]) {
      var row = existingRow[tw];
      StreamersSheet.getRange(row, 2).setValue(info.discordUsername); // keep display name current
      StreamersSheet.getRange(row, 5).setValue(now); // Last Seen
    } else {
      StreamersSheet.appendRow([info.discordUserId, info.discordUsername, tw, now, now]);
    }
  });

  return usernames.length;
}


function getTwitchAppToken() {
  var props = PropertiesService.getScriptProperties();
  var cached = props.getProperty('TWITCH_TOKEN_CACHE');
  if (cached) {
    try {
      var c = JSON.parse(cached);
      if (c.expiresAt > Date.now() + 60000) return c.token;
    } catch (err) { /* fall through and re-fetch */ }
  }

  var clientId = props.getProperty('TWITCH_CLIENT_ID');
  var clientSecret = props.getProperty('TWITCH_CLIENT_SECRET');
  if (!clientId || !clientSecret) {
    console.error('getTwitchAppToken: TWITCH_CLIENT_ID or TWITCH_CLIENT_SECRET not set in Script Properties.');
    return null;
  }

  var resp = UrlFetchApp.fetch('https://id.twitch.tv/oauth2/token', {
    method: 'post',
    payload: {
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: 'client_credentials'
    },
    muteHttpExceptions: true
  });

  var data;
  try {
    data = JSON.parse(resp.getContentText());
  } catch (err) {
    console.error('getTwitchAppToken: could not parse Twitch auth response: ' + err);
    return null;
  }
  if (!data.access_token) {
    console.error('getTwitchAppToken: Twitch auth failed: ' + resp.getContentText());
    return null;
  }

  props.setProperty('TWITCH_TOKEN_CACHE', JSON.stringify({
    token: data.access_token,
    expiresAt: Date.now() + (data.expires_in * 1000)
  }));
  return data.access_token;
}


function getLiveStreamsBatch(usernames) {
  if (!usernames.length) return [];

  var token = getTwitchAppToken();
  if (!token) return [];
  var clientId = PropertiesService.getScriptProperties().getProperty('TWITCH_CLIENT_ID');

  var query = usernames.map(function (u) { return 'user_login=' + encodeURIComponent(u); }).join('&');
  var resp = UrlFetchApp.fetch('https://api.twitch.tv/helix/streams?' + query, {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Client-Id': clientId
    },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    console.error('getLiveStreamsBatch: Twitch streams fetch failed (' + resp.getResponseCode() + '): ' + resp.getContentText());
    return [];
  }

  try {
    var data = JSON.parse(resp.getContentText());
    return data.data || [];
  } catch (err) {
    console.error('getLiveStreamsBatch: could not parse Twitch response: ' + err);
    return [];
  }
}


function getStreamsData() {
  var now = new Date().getTime();
  if (streamsCache && streamsCacheTime && (now - streamsCacheTime) < STREAMS_CACHE_DURATION) {
    return streamsCache;
  }

  try {
    if (!StreamersSheet) {
      return {
        success: false,
        error: 'No "Streamers" sheet found.',
        message: 'Streams feature is not set up yet.'
      };
    }

    var data = StreamersSheet.getDataRange().getValues();
    var recencyCutoff = now - (STREAM_RECENCY_HOURS * 60 * 60 * 1000);
    var streamers = [];
    for (var i = 1; i < data.length; i++) {
      var tw = (data[i][2] || '').toString().trim();
      if (!tw) continue;

      
      var lastSeenRaw = data[i][4];
      var lastSeenMs = lastSeenRaw instanceof Date
        ? lastSeenRaw.getTime()
        : (lastSeenRaw ? new Date(lastSeenRaw).getTime() : NaN);
      if (isNaN(lastSeenMs) || lastSeenMs < recencyCutoff) continue;

      streamers.push({
        discordUsername: data[i][1] || 'Unknown',
        twitchUsername: tw
      });
    }

    var byLoginLower = {};
    streamers.forEach(function (s) { byLoginLower[s.twitchUsername.toLowerCase()] = s; });

    var loginList = streamers.map(function (s) { return s.twitchUsername.toLowerCase(); });
    var liveByLogin = {};
    // Twitch caps user_login params at 100 per request, so batch if needed.
    for (var i = 0; i < loginList.length; i += 100) {
      var batch = loginList.slice(i, i + 100);
      getLiveStreamsBatch(batch).forEach(function (s) {
        liveByLogin[s.user_login.toLowerCase()] = s;
      });
    }

    var live = [];
    Object.keys(liveByLogin).forEach(function (loginLower) {
      var stream = liveByLogin[loginLower];
      var known = byLoginLower[loginLower];
      live.push({
        twitchUsername: known ? known.twitchUsername : stream.user_name,
        discordUsername: known ? known.discordUsername : '',
        title: stream.title || '',
        game: stream.game_name || '',
        viewers: stream.viewer_count || 0,
        thumbnail: (stream.thumbnail_url || '').replace('{width}', '440').replace('{height}', '248'),
        startedAt: stream.started_at || ''
      });
    });
    live.sort(function (a, b) { return b.viewers - a.viewers; });

    var result = {
      success: true,
      live: live,
      knownStreamerCount: streamers.length,
      lastUpdated: new Date().toISOString()
    };

    streamsCache = result;
    streamsCacheTime = now;
    return result;

  } catch (error) {
    console.error('Error in getStreamsData:', error);
    return {
      success: false,
      error: error.toString(),
      message: 'Failed to load streams'
    };
  }
}

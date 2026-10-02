var App = SpreadsheetApp;
var SS = App.getActiveSpreadsheet();
var ListPage = SS.getSheetByName("List");
var PIDTable = SS.getSheetByName("PID Conversion Table");
var PlayerClearsLocal = SS.getSheetByName("Player Clears Local Storage");

var cachedResponse = null;
var cacheTime = null;
var CACHE_DURATION = 300000; // 5 minutes in milliseconds

// MAIN DOGET FUNCTION (UPDATED)

function doGet(e) {
  console.log("doGet called, action:", e.parameter ? e.parameter.action : "none");
  var action = e.parameter ? e.parameter.action : null;
  var placement = e.parameter ? e.parameter.placement : null;
  var callback = e.parameter ? e.parameter.callback : null;
  var pathInfo = e.pathInfo;

  // Handle robots.txt
  if (pathInfo === "robots.txt") {
    var robotsTxt = "User-agent: *\n";
    robotsTxt += "Allow: /\n";
    robotsTxt += "Sitemap: " + ScriptApp.getService().getUrl() + "?action=sitemap\n";

    return ContentService.createTextOutput(robotsTxt)
      .setMimeType(ContentService.MimeType.TEXT);
  }

  // Handle sitemap
  if (action === "sitemap") {
    return generateSitemap();
  }

  // Handle existing API endpoints
  if (action === "getByPlacementLocal" && placement) {
    return getByPlacementLocal(e, placement);
  }

  if (action === "getAllLevels") {
  var result = getAllLevelsCACHED();
  var jsonOutput = JSON.stringify(result);
  if (callback) {
    jsonOutput = callback + "(" + jsonOutput + ");";
    return ContentService.createTextOutput(jsonOutput)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonOutput)
    .setMimeType(ContentService.MimeType.JSON);
}
  if (action === "getLeaderboard") {
    console.log("PID table rows:", PIDTable.getLastRow());
    var result = getLeaderboardData();
    var jsonOutput = JSON.stringify(result);
    if (callback) {
      jsonOutput = callback + "(" + jsonOutput + ");";
      return ContentService.createTextOutput(jsonOutput)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(jsonOutput)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // getStreamsData() lives in Streams.gs
  // polling
  if (action === "getStreams") {
    var result = getStreamsData();
    var jsonOutput = JSON.stringify(result);
    if (callback) {
      jsonOutput = callback + "(" + jsonOutput + ");";
      return ContentService.createTextOutput(jsonOutput)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(jsonOutput)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // For everything else
  return getHTMLPage();
}

function getHTMLPage() {
  var htmlContent = `
  <!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Mario Maker 2 Leaderboard - Top 100 Hardest Levels</title>
    <meta name="description" content="The official Mario Maker 2 leaderboard tracking the top 100 hardest levels. See rankings, clear counts, and who has beaten each level.">
    <meta name="keywords" content="Mario Maker 2, MM2, leaderboard, top 100, hardest levels, Mario Maker rankings">
    <meta name="robots" content="index, follow">
    <link rel="canonical" href="${ScriptApp.getService().getUrl()}">

    <!-- Open Graph tags for social media -->
    <meta property="og:title" content="Mario Maker 2 Leaderboard">
    <meta property="og:description" content="Track the top 100 hardest Mario Maker 2 levels and see who has cleared them.">
    <meta property="og:type" content="website">
    <meta property="og:url" content="${ScriptApp.getService().getUrl()}">

    <style>
      body {
        font-family: Arial, sans-serif;
        max-width: 1200px;
        margin: 0 auto;
        padding: 20px;
        background-color: #f5f5f5;
      }
      h1 {
        color: #e60012;
        text-align: center;
      }
      .level-list {
        background: white;
        border-radius: 8px;
        padding: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      }
      .level-item {
        border-bottom: 1px solid #eee;
        padding: 10px;
        display: flex;
        justify-content: space-between;
      }
      .level-rank {
        font-weight: bold;
        color: #e60012;
        width: 60px;
      }
      .level-name {
        flex: 1;
      }
      .level-clears {
        width: 80px;
        text-align: right;
      }
      .loading {
        text-align: center;
        padding: 40px;
        color: #666;
      }
      footer {
        text-align: center;
        margin-top: 40px;
        padding: 20px;
        color: #666;
        font-size: 12px;
      }
    </style>
  </head>
  <body>
    <h1>🏆 Mario Maker 2 Leaderboard 🏆</h1>
    <p style="text-align: center;">The top 100 hardest levels in Super Mario Maker 2, ranked by difficulty and clear rate.</p>

    <div class="level-list">
      <div class="loading">Loading leaderboard data...</div>
    </div>

    <footer>
      <p>Data updates automatically every 12 hours. Last updated: ${new Date().toLocaleString()}</p>
      <p>This leaderboard tracks clears from the community and ranks levels based on difficulty.</p>
    </footer>

    <script>
      // Fetch the actual data after page loads
      fetch('?action=getAllLevels')
        .then(response => response.json())
        .then(data => {
          const container = document.querySelector('.level-list');
          if (data && data.levels) {
            let html = '<div style="font-weight: bold; display: flex; padding: 10px; background: #f0f0f0;">';
            html += '<div style="width: 60px;">Rank</div>';
            html += '<div style="flex: 1;">Level Name</div>';
            html += '<div style="width: 80px; text-align: right;">Clears</div>';
            html += '</div>';

            data.levels.forEach(level => {
              html += '<div class="level-item">';
              html += '<div class="level-rank">#' + level.placement + '</div>';
              html += '<div class="level-name">' + escapeHtml(level.name) + '</div>';
              html += '<div class="level-clears">' + level.clears + ' clears</div>';
              html += '</div>';
            });

            container.innerHTML = html;
          } else {
            container.innerHTML = '<p>No data available</p>';
          }
        })
        .catch(error => {
          console.error('Error loading data:', error);
          document.querySelector('.level-list').innerHTML = '<p>Error loading leaderboard data. Please try again later.</p>';
        });

      function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
      }
    </script>
  </body>
  </html>
  `;

  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle("Mario Maker 2 Leaderboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAllLevelsForAPI(callback) {
  var levels = [];

  for (var i = 1; i <= 100; i++) {
    try {
      var row = placementToRow(i);
      var clearCount = 0;
      var col = 4;

      while (PlayerClearsLocal.getRange(row, col).getValue() != "") {
        clearCount++;
        col++;
      }

      levels.push({
        placement: i,
        name: ListPage.getRange(i, 1).getValue() || "Unknown",
        courseID: PlayerClearsLocal.getRange(row, 1).getValue() || "DELETED",
        uploader: PlayerClearsLocal.getRange(row, 3).getValue() || "Unknown",
        clears: clearCount
      });
    } catch (error) {
      console.error("Error getting level", i, ":", error);
      levels.push({
        placement: i,
        name: "Level " + i,
        courseID: "ERROR",
        uploader: "Unknown",
        clears: 0
      });
    }
  }

  var result = {
    total: 100,
    levels: levels,
    lastUpdated: new Date().toISOString(),
    source: "Mario Maker 2 Community Leaderboard"
  };

  var jsonOutput = JSON.stringify(result);

  // Handle JSONP callback if present (for compatibility with existing site)
  if (callback) {
    jsonOutput = callback + "(" + jsonOutput + ");";
    return ContentService.createTextOutput(jsonOutput)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(jsonOutput)
    .setMimeType(ContentService.MimeType.JSON);
}

function generateSitemap() {
  var baseUrl = ScriptApp.getService().getUrl();
  var sitemap = '<?xml version="1.0" encoding="UTF-8"?>\n';
  sitemap += '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n';
  sitemap += '  <url>\n';
  sitemap += '    <loc>' + baseUrl + '</loc>\n';
  sitemap += '    <lastmod>' + new Date().toISOString().split('T')[0] + '</lastmod>\n';
  sitemap += '    <changefreq>daily</changefreq>\n';
  sitemap += '    <priority>1.0</priority>\n';
  sitemap += '  </url>\n';
  sitemap += '</urlset>';

  return ContentService.createTextOutput(sitemap)
    .setMimeType(ContentService.MimeType.XML);
}

// CACHED VERSION OF GETALLLEVELS yep
function getAllLevelsCACHED() {
  var now = new Date().getTime();

  // Return cached response if valid
  if (cachedResponse && cacheTime && (now - cacheTime) < CACHE_DURATION) {
    console.log("Returning cached response");
    return cachedResponse;
  }

  console.log("Fetching fresh data...");
  var result = getAllLevelsFAST();

  // Cache the result
  cachedResponse = result;
  cacheTime = now;

  return result;
}

 
// FAST VERSION wee

function getAllLevelsFAST() {
  console.log("getAllLevelsFAST started");
  var startTime = new Date();

  try {
    // Get List data (100 levels)
    var listData = ListPage.getRange("A1:H100").getValues();

    // Get PlayerClearsLocal data
    var clearsData = [];
    try {
      clearsData = PlayerClearsLocal.getDataRange().getValues();
    } catch (err) {
      console.log("Warning: Could not read PlayerClearsLocal:", err);
    }

    // Create placement to row mapping
    var placementMap = {};
    for (var i = 0; i < clearsData.length; i++) {
      if (clearsData[i][1]) { // Column B has placement
        placementMap[clearsData[i][1]] = i;
      }
    }

    // 4. Process all levels
    var levels = [];
    for (var i = 0; i < 100; i++) {
      var listRow = listData[i];

      // Basic info from List sheet
      var placement = i + 1;
      var courseName = listRow[0] || "Unknown";
      var clearVideo = listRow[1] || "";
      var courseID = listRow[2] || "XXX-XXX-XXX";
      var uploaderCell = listRow[3] || "";
      var points = listRow[4] || 0;
      var genre = listRow[5] || "Unknown";
      var originalCourseID = listRow[7] || ""; // Column H

      // Parse creator from uploader cell
      var creator = uploaderCell;
      if (uploaderCell && uploaderCell.includes("(")) {
        creator = uploaderCell.slice(uploaderCell.indexOf("(") + 1, uploaderCell.indexOf(")"));
      }

      // Get clears and uploader info
      var clearsCount = 0;
      var uploaderContent = "";

      if (placementMap[placement] !== undefined) {
        var clearsRowIndex = placementMap[placement];
        var clearsRow = clearsData[clearsRowIndex];

        // Uploader is in column C (index 2)
        uploaderContent = clearsRow[2] || "";

        // Count clears from column D onward (index 3+)
        var col = 2;
        while (col < clearsRow.length && clearsRow[col] && clearsRow[col] !== "") {
          clearsCount++;
          col++;
        }
      }

      // Parse uploader info
      var uploaderInfo = parsePlayerContent(uploaderContent);

      // Add level to array
      levels.push({
        Placement: placement,
        CourseName: courseName,
        ClearVideo: clearVideo,
        CourseID: courseID,
        Uploader: uploaderInfo.displayName,
        UploaderPID: uploaderInfo.pid,
        UploaderRegion: uploaderInfo.region,
        Creator: creator,
        Points: Number(points),
        Genre: genre,
        Clears: clearsCount,
        OriginalCourseID: originalCourseID,
      });
    }

    var endTime = new Date();
    var timeTaken = endTime - startTime;
    console.log("Processed " + levels.length + " levels in " + timeTaken + "ms");

    // Return SUCCESS response
    return {
      success: true,
      count: levels.length,
      lastUpdated: new Date().toISOString(),
      processingTime: timeTaken + "ms",
      levels: levels
    };

  } catch (error) {
    console.error("Error in getAllLevelsFAST:", error);

    // Return ERROR response
    return {
      success: false,
      error: error.toString(),
      message: "Failed to load levels"
    };
  }
}

// LEADERBOARD FUNCTIONS

function getLeaderboardData() {
  console.log("getLeaderboardData started");
  var startTime = new Date();

  try {
    // 1. Get PID Table data
    var pidData = PIDTable.getRange("A1:C" + PIDTable.getLastRow()).getValues();

    // 2. Get PlayerClearsLocal data for counting clears
    var clearsData = [];
    try {
      clearsData = PlayerClearsLocal.getDataRange().getValues();
    } catch (err) {
      console.log("Warning: Could not read PlayerClearsLocal for clear counting:", err);
    }

    // 3. Process players
    var players = [];

    for (var i = 0; i < pidData.length; i++) { // Start from 1 to skip header

      var pid = pidData[i][0]; // Column A: PID
      console.log("Row", i, "pid:", pid, "type:", typeof pid);
      var playerName = pidData[i][1] || ""; // Column B: Player Name
      var points = pidData[i][2] || 0; // Column C: Points

      // Skip invalid/empty PIDs
      if (!pid || pid === "" || pid === 0) continue;

      // Parse player info from the name format "XX Name PID"
      var playerInfo = parsePlayerContent(playerName);

      // Count clears for this player, and get which placements they cleared /
      // uploaded so player profile pages can show a full breakdown.
      var clearData = getPlayerClearData(pid, clearsData);

      // Get rank based on points
      var rank = getRankFromPoints(points);

      players.push({
        pid: pid.toString(),
        name: playerInfo.name,
        region: playerInfo.region,
        points: Number(points),
        clears: clearData.clears,
        clearPlacements: clearData.clearPlacements,
        uploadedPlacements: clearData.uploadedPlacements,
        rank: rank,
        displayName: playerInfo.displayName
      });
    }

    // sort by points descending
    console.log("Total players before filter:", players.length);
    players = players.filter(function(p) { return p.points > 0; });
    console.log("Total players after filter:", players.length);

    players.sort((a, b) => b.points - a.points);

    var endTime = new Date();
    var timeTaken = endTime - startTime;
    console.log("Processed " + players.length + " players in " + timeTaken + "ms");

    return {
      success: true,
      count: players.length,
      lastUpdated: new Date().toISOString(),
      processingTime: timeTaken + "ms",
      players: players
    };

  } catch (error) {
    console.error("Error in getLeaderboardData:", error);

    return {
      success: false,
      error: error.toString(),
      message: "Failed to load leaderboard data"
    };
  }
}

// Returns which top 100 placements a player cleared, and which of those
// placements they uploaded. Mirrors the exact matching logic of
// countPlayerClears (same start column, same "PID appears anywhere in the
// row from column C onward" rule) so clearPlacements.length always equals
// what countPlayerClears would have returned
function getPlayerClearData(pid, clearsData) {
  var pidStr = pid.toString();
  var clearPlacements = [];
  var uploadedPlacements = [];

  for (var row = 0; row < clearsData.length; row++) {
    var placement = clearsData[row][1];
    // Only consider rows with a valid placement 1-100 (same guard as countPlayerClears)
    if (placement === "" || placement === null || placement === undefined ||
        placement === "#N/A" || isNaN(placement) ||
        placement < 1 || placement > 100) {
      continue;
    }

    // Does this player's PID appear anywhere in the row from column C (index 2) onward?
    var matchedThisRow = false;
    for (var col = 2; col < clearsData[row].length; col++) {
      var cellValue = clearsData[row][col];
      if (cellValue && cellValue.toString().includes(pidStr)) {
        matchedThisRow = true;
        break;
      }
    }
    if (matchedThisRow) {
      clearPlacements.push(placement);
    }

    // Uploader is specifically column C (index 2)
    var uploaderCell = clearsData[row][2];
    if (uploaderCell && uploaderCell.toString().includes(pidStr)) {
      uploadedPlacements.push(placement);
    }
  }

  return {
    clearPlacements: clearPlacements,
    uploadedPlacements: uploadedPlacements,
    clears: clearPlacements.length
  };
}

function countPlayerClears(pid, clearsData) {
  var count = 0;
  var pidStr = pid.toString();

  for (var row = 0; row < clearsData.length; row++) {
    var placement = clearsData[row][1];
    // Only count rows with a valid placement 1-100
     if (placement === "" || placement === null || placement === undefined ||
        placement === "#N/A" || isNaN(placement) ||
        placement < 1 || placement > 100) {
      continue;
    }

    for (var col = 2; col < clearsData[row].length; col++) {
      var cellValue = clearsData[row][col];
      if (cellValue && cellValue.toString().includes(pidStr)) {
        count++;
        break;
      }
    }
  }

  return count;
}

function getRankFromPoints(points) {
  points = Number(points);

  if (points >= 5000) return "GOAT";
  if (points >= 2500) return "Legend";
  if (points >= 1000) return "Grandmaster";
  if (points >= 500) return "Master";
  if (points >= 250) return "Professional";
  if (points >= 100) return "Advanced";
  if (points >= 50) return "Intermediate";
  if (points >= 10) return "Novice";
  return "Unranked";
}


// HELPER FUNCTIONS

function parsePlayerContent(content) {
  if (!content || content.trim() === "") {
    return {
      region: "XX",
      name: "Unknown",
      pid: "0",
      displayName: "XX Unknown 0"
    };
  }

  // Format: "XX Name PID" or "XX Name Other PID"
  var parts = content.trim().split(" ");

  if (parts.length < 3) {
    return {
      region: "XX",
      name: content || "Unknown",
      pid: "0",
      displayName: content || "XX Unknown 0"
    };
  }

  var region = parts[0];
  var pid = parts[parts.length - 1];
  var name = parts.slice(1, parts.length - 1).join(" ");

  return {
    region: region,
    name: name,
    pid: pid,
    displayName: content
  };
}


// APP INFO FUNCTION

function getAppInfo() {
  return {
    success: true,
    app: "SMM2 TOP100 API",
    version: "3.0",
    description: "API for SMM2 TOP100 levels and leaderboard",
    endpoints: [
      "/getAllLevels",
      "/getLeaderboard",
      "/getStreams",
      "/getAppInfo"
    ],
    timestamp: new Date().toISOString()
  };
}


// LEGACY FUNCTIONS (Keep for compatibility)

function getByPlacementLocal(e, placement) {
  // Keep this function if other code calls it
  try {
    var levelData = getLevelData(parseInt(placement));
    return {
      success: true,
      level: levelData
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getLevelData(placement) {
  // Simplified version for single level
  var listRow = ListPage.getRange(placement, 1, 1, 6).getValues()[0];

  return {
    Placement: placement,
    CourseName: listRow[0] || "Unknown",
    ClearVideo: listRow[1] || "",
    CourseID: listRow[2] || "XXX-XXX-XXX",
    Creator: listRow[3] || "Unknown",
    Points: listRow[4] || 0,
    Genre: listRow[5] || "Unknown"
  };
}

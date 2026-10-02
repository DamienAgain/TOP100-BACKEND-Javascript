function testLocalBackend() {
  //updatePlacementsRange("Range", 6, 7)

  //local functions
  //findNewLevels() //HAS TO FINISH EXECUTION BEFORE NEXT FUNCTION CAN RUN
  //deletedLevelsHandler()

  //api functions
  //requestAllListUploaders()
  //updatePlacementsRange("Range", 80, 100)
  leaderboardUpdatev2()
}

async function apiReqest(url, localData) {
  const params = { "method": "get", "muteHttpExceptions": true };
  try {
    const response = UrlFetchApp.fetch(url, params);
    const rawText = response.getContentText()
    var safeText = rawText
        .replace(/"pid":(\d+)/g, '"pid":"$1"')
        .replace(/"cleared":\[([^\]]+)\]/g, function(match, nums) {
          return '"cleared":[' + nums.split(',').map(function(n) { return '"' + n.trim() + '"'; }).join(',') + ']';
        })
        .replace(/"liked":\[([^\]]+)\]/g, function(match, nums) {
          return '"liked":[' + nums.split(',').map(function(n) { return '"' + n.trim() + '"'; }).join(',') + ']';
  })
    const json = JSON.parse(safeText)

    if (json == "") {
      console.log("the api didnt send anything, exiting command")
      throw new SyntaxError("Empty json")
    } else {
      //run the list and sheet code here inside the async func heh
      parseJsonData(json, localData);
    }

  } catch (error) {
    console.error(error);
  }
}

function updatePlacementsRange(type, startIndex, endIndex) {

  async function levelPlayedRequest(url, courseRow) {
    const params = { "method": "get", "muteHttpExceptions": true };

    try {
      const response = UrlFetchApp.fetch(url, params);
      if (response.getResponseCode() == 200) {
        const rawText = response.getContentText()
        var safeText = rawText
        .replace(/"pid":(\d+)/g, '"pid":"$1"')
        .replace(/"cleared":\[([^\]]+)\]/g, function(match, nums) {
        return '"cleared":[' + nums.split(',').map(function(n) { return '"' + n.trim() + '"'; }).join(',') + ']';
        })
        .replace(/"liked":\[([^\]]+)\]/g, function(match, nums) {
        return '"liked":[' + nums.split(',').map(function(n) { return '"' + n.trim() + '"'; }).join(',') + ']';
  })
        const json = JSON.parse(safeText)
        //Run The List And Sheet Code Here Inside The Async Func
        levelPlayedParse(json, courseRow);

      } else if (response.getResponseCode() == 400) {
        console.log(response.getResponseCode(), "failed to get a server response just gonna skip this cause its probably deleted")
        placementIndex++
        throw new SyntaxError(response)

      } else if (response.getResponseCode() == 502) {
        console.log(response.getResponseCode(), "bad gateway, skipping this level")
        placementIndex++
        throw new SyntaxError(response)
      } else {
        console.log(response.getResponseCode())
        console.log("sleeping for 5000 due to this error")
        Utilities.sleep(5100)
        throw new SyntaxError("Empty json")
      }


    } catch (error) {
      console.error(error);
      //retry that entry
      placementIndex -= 1
    }
  }
  //type Handler
  var localData = getAllValidCourseIDs()

  if (type == "typeA") {
    console.log("type A")
    startIndex = 0
    endIndex = 47
  } else if (type == "typeB") {
    console.log("type B")
    startIndex = 47
    endIndex = localData.placements.length
  } else if (type == "Range") {
    console.log("range from " + startIndex + " to " + endIndex)
  }
  for (placementIndex = startIndex; placementIndex < endIndex; placementIndex++) {
    var courseRow = placementToRow(localData.placements[placementIndex])
    var courseID = PlayerClearsLocal.getRange(courseRow, 1).getValue()
    var url = "https://tgrcode.com/mm2/level_played/" + courseID
    console.log("~checking placement~ " + localData.placements[placementIndex])
    levelPlayedRequest(url, courseRow)
    Utilities.sleep(1000)
  }
}


function changeLogClear(username, placement) {
  var title = ListPage.getRange(placement, 1).getValue()

  console.log("Clear by '" + username + "' added on '" + title + "'(" + (placement) + ") on Date: " + formatedDate)
  ChangeLog.getRange(ChangeLog.getLastRow() + 1, 1).setValue("Clear by '" + username + "' added on '" + title + "' on Date: " + formatedDate)
}

function changeLogAdd(levelTitle, placement, additions) {
  var listIDs = PlayerClearsLocal.getRange(1, 1, 100).getValues().map(i => i[0])
  var placements = PlayerClearsLocal.getRange(1, 2, 100).getValues().map(i => i[0])
  //console.log(listIDs[placements.filter(Number).length + additions])
  var pushedlevel = listIDs[placements.filter(Number).length + additions]

  console.log(levelTitle + " added at placement #" + placement + ", this pushes '" + pushedlevel + "' off. Date: " + formatedDate)
  ChangeLog.getRange(ChangeLog.getLastRow() + 1, 1).setValue(levelTitle + " added at placement #" + placement + ", this pushes '" + pushedlevel + "' off. Date: " + formatedDate)
}


function levelPlayedParse(json, courseRow) {
  //check if the level has any exceptions
  var exceptionsArray = Exceptions.getRange(1, 1, Exceptions.getLastRow()).getValues().map(function (row) { return row[0]; })
  var exceptionsRow = exceptionsArray.indexOf(PlayerClearsLocal.getRange(courseRow, 1).getValue()) + 1

  //find all current clears on the sheet
  var columnCounter = 4
  var currentClearsPIDs = ""
  while (PlayerClearsLocal.getRange(courseRow, columnCounter).getValue() != "") {
    currentClearsPIDs += extractThePID(PlayerClearsLocal.getRange(courseRow, columnCounter).getValue()) + " "
    columnCounter++
  }

  //check for any new clears as PIDs
  var newClears = []
  if (json.cleared != undefined) {
  for (index = 0; index < json.cleared.length; index++) {
    var pid = (json.cleared[index] + "").slice(0, -5)

    if (!currentClearsPIDs.includes(pid)) {
      if (xExceptionsHandler(exceptionsRow, pid) == true) { continue }
      if (invalidPIDList.includes(parseInt(pid))) { continue }

      console.log("new clear found for placement " + courseRow)
      // pass the full player object so pidToUserContent can use it
      newClears.push(pidToUserContent(json.cleared[index] + "", json.players))
    }
  }
}
  //add any new extra clears
  newClears = newClears.concat(oExceptionHandler(exceptionsRow, currentClearsPIDs))

  //if theres nothing new we can exit the function 
  if (newClears.length == 0) {
    return;
  }

  //append all the new clears to the sheet
  console.log("printing new clears: " + newClears)
  for (index = 0; index < newClears.length; index++) {
    changeLogClear(newClears[index], courseRow)

    PlayerClearsLocal.getRange(courseRow, columnCounter + index).setValue(newClears[index])
  }
}

function xExceptionsHandler(exceptionsRow, pid) {
  //check if there might be an exception (only Xs should be checked here due to the restrictions)
  if (exceptionsRow > 0) {
    var concatTrick = ""
    //use concat trick for exceptionsEntries so we can search with indluces method
    Exceptions.getRange(exceptionsRow, 2, 1, Exceptions.getLastColumn()).getValues().forEach((i) => concatTrick += i + " ")
    if (concatTrick.includes(pid)) {
      return true;
    }
  }
  return false
}

function oExceptionHandler(exceptionsRow, currentClearsPIDs) {
  if (exceptionsRow <= 0) {
    return [];
  }
  var exceptionsEntries = Exceptions.getRange(exceptionsRow, 2, 1, Exceptions.getLastColumn()).getValues()

  var endsWithO = exceptionsEntries[0].map(i => i.endsWith("O"))
  var exceptionsPIDList = exceptionsEntries[0].map(i => extractThePID(i.slice(0, -2)))
  var newClears = []
  for (index = 0; index < exceptionsEntries[0].length; index++) {
    if (!currentClearsPIDs.includes(exceptionsPIDList[index] + "") && endsWithO[index]) {
      console.log("player found with valid exception clear not currently on the lsit " + exceptionsEntries[0][index].slice(0, -2))
      newClears.push(exceptionsEntries[0][index].slice(0, -2))
    }
  }
  return newClears;
}

function pidToUserContent(pid, playersArrayJSON) {

  //first, check the PID conversion table for the PID
  var pidTable = PIDTable.getRange(1, 1, PIDTable.getLastRow()).getValues().map(function (row) { return row[0]; });
  var row = pidTable.indexOf(Number(pid.slice(0, -5))) + 1
  if (row != 0) {
    return PIDTable.getRange(row, 2).getValue() + pid.slice(0, -5)
  }

  //next check the level data in the json package
  for (index = 0; index < playersArrayJSON.length; index++) {
    if (pid == playersArrayJSON[index].pid) {
      return playersArrayJSON[index].country + " " + playersArrayJSON[index].name + " " + pid.slice(0, -5)
    }
  }

  //all else failed so we return XX null
  return "XX Null " + pid.slice(0, -5)
}

function getAllValidCourseIDs() {
  var deletedLevelsPlacements = [];
  var validIDs = "";
  var placements = [];

  for (index = 1; index <= 100; index++) {
    if (ListPage.getRange(index, 3).getValue()[1] != " ") {
      validIDs += ListPage.getRange(index, 3).getValue() + ",";
      placements.push(index);
    } else {
      deletedLevelsPlacements.push(index);
    }
  }
  var localData = { "ids": validIDs.slice(0, -1), "placements": placements, "deletedPlacements": deletedLevelsPlacements }
  return localData;
}

function requestAllListUploaders() {
  var localData = getAllValidCourseIDs();
  var url = "https://tgrcode.com/mm2/level_info_multiple/" + localData.ids;
  apiReqest(url, localData)
}

function parseJsonData(json, localData) {
  //Update for uploaders of non deleted levels
  //console.log("updating uploaders");
  for (index = 0; index < json.courses.length; index++) {
    var uploaderInfo = json.courses[index].uploader.country + " " + json.courses[index].uploader.name + " " + (json.courses[index].uploader.pid).toString().slice(0, -5)
    printUploaderInfoLocal(uploaderInfo, localData.placements[index]);
  }
}

function extractThePID(userContent) {
  pid = ""
  for (var i = 1; i < userContent.length; i++) {
    pid = userContent.slice(userContent.length - i, userContent.length);

    if (pid.includes(" ")) {
      return pid.slice(1)
    }
  }
  return pid
}

function deletedLevelsHandler() {
  PlayerClearsLocal.sort(2)

  deletedLevelsArray = DeletedLevelData.getRange(1, 1, DeletedLevelData.getLastRow()).getValues().map(function (row) { return row[0]; })
  localData = getAllValidCourseIDs()

  for (index = 0; index < localData.deletedPlacements.length; index++) {
    //console.log(deletedLevelsArray.indexOf(PlayerClearsLocal.getRange(localData.deletedPlacements[index], 1).getValue()), localData.deletedPlacements[index])

    printDeletedInfoLocal(deletedLevelsArray.indexOf(PlayerClearsLocal.getRange(localData.deletedPlacements[index], 1).getValue()), localData.deletedPlacements[index])
  }

}

function printUploaderInfoLocal(uploaderInfo, placement) {
  //since the PlayerClearsLocal sheet isnt ordered by placement, we have to match placements to their actual row on the sheet 
  if (placement != PlayerClearsLocal.getRange(placement, 3)) {
    var row = placementToRow(placement)
  } else {
    var row = placement
  }

  if (PlayerClearsLocal.getRange(row, 3).getValue() == "") {
    PlayerClearsLocal.getRange(row, 3).setValue(uploaderInfo);
    return;

  } else if (PlayerClearsLocal.getRange(row, 3).getValue() == uploaderInfo) {
    return;

  } else {
    console.log("uploaders do not match, we not good on D: " + row)
    //check to make sure the PIDs are the same, if so then the person must have changed their name so we can change that safely

    if (extractThePID(uploaderInfo) == extractThePID(PlayerClearsLocal.getRange(row, 3).getValue())) {
      console.log("PIDs match, updating user name and or country")
      PlayerClearsLocal.getRange(row, 3).setValue(uploaderInfo)
    } else {
      console.log("PIDs dont match, somethings gone wrong... D: this not good")
    }
    return;
  }
}

function printDeletedInfoLocal(index, placement) {
  if (index < 0 || index === null || index === undefined) {
    console.log("No deleted level data found for placement " + placement + ", skipping")
    return;
  }
  
  if (placement != PlayerClearsLocal.getRange(placement, 3)) {
    var row = placementToRow(placement)
  } else {
    var row = placement
  }

  var deletedCourseInfo = DeletedLevelData.getRange(index + 1, 2, 1, DeletedLevelData.getLastColumn() - 1).getValues()
  PlayerClearsLocal.getRange(row, 3, 1, deletedCourseInfo[0].length).setValues(deletedCourseInfo)
}

function placementToRow(placement) {
  var placements = PlayerClearsLocal.getRange(1, 2, PlayerClearsLocal.getLastRow()).getValues()
  for (var i = 0; i < placements.length; i++) {
    if (placements[i][0] == placement) {
      return i + 1
    }
  }
  console.error("placementToRow: could not find placement " + placement)
  return null
}

function findNewLevels() {
  var lastRow = PlayerClearsLocal.getLastRow() + 1
  var localIDs = PlayerClearsLocal.getRange("A1:A" + lastRow).getValues().map(function (row) { return row[0]; });
  var existingPlacements = PlayerClearsLocal.getRange("B1:B" + lastRow).getValues().map(function (row) { return row[0]; });
  var listIDs = ListPage.getRange("C1:C100").getValues().map(function (row) { return row[0]; });
  var newLevels = 0

  for (index = 0; index < 100; index++) {
    if (localIDs.indexOf(listIDs[index]) == -1) {
      var placement = index + 1;
      if (existingPlacements.indexOf(placement) != -1) {
        console.log("placement " + placement + " already exists in column B, skipping to avoid duplicate");
        continue;
      }
      console.log("level not on local clears page found at placement ", placement, " adding it now");
      changeLogAdd(ListPage.getRange(placement, 1).getValue(), placement, newLevels)
      PlayerClearsLocal.getRange(lastRow, 1).setValue(ListPage.getRange(placement, 3).getValue())
      PlayerClearsLocal.getRange(lastRow, 2).setValue(placement)
      newLevels += 1
      lastRow += 1
    }
  }
  //sort PlayerClearsLocal for consistancy, sorting doesnt even work cause we need to wait for the trigger to end before we can even sort lol
  PlayerClearsLocal.sort(2)
}

function getLevelPlayedInfo(placement) {
  var courseRow = placementToRow(placement)
  var courseID = PlayerClearsLocal.getRange(courseRow, 1).getValue()
  var url = "https://tgrcode.com/mm2/level_played/" + courseID
  levelPlayedRequest(url, courseRow)
}

function createListUpdateTrigger() {
  ByPass.getRange("A1:A5").setValue("")
  ByPass.getRange("B1:B5").setValue("0")
  ByPass.getRange(1, 1).setValue("true")
  ScriptApp.newTrigger("updateListEveryMin")
    .timeBased()
    .everyMinutes(1)
    .create();

  var triggers = getProjectTriggersByName('createListUpdateTrigger');
  for (var i = 0; i < triggers.length; ++i) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function restart() {
  var minTriggers = getProjectTriggersByName('updateListEveryMin');
  var createTriggers = getProjectTriggersByName('createListUpdateTrigger');
  if (minTriggers.length == 0 && createTriggers.length == 0) {
    console.log("error found, resetting triggers");
    createListUpdateTrigger();
  } else {
    console.log("program is healthy :)");
  }
}

function failSafe(index) {
  if (ByPass.getRange(index, 2).getValue() > 6) {
    ByPass.getRange(index, 2).setValue(0)
    ByPass.getRange(index, 1).setValue("true")
  } else if (ByPass.getRange(index, 2).getValue() <= 1) {
    ByPass.getRange(index, 2).setValue(1)
  }
}

function updateListEveryMin() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    console.log("Another instance is running, skipping this trigger execution");
    return;
  }
  
  try {
  var lastIndex = 0
  for (var index = 1; index <= 7; index++) {
    if (ByPass.getRange(index, 2).getValue() > 0) {
      ByPass.getRange(index, 2).setValue(ByPass.getRange(index, 2).getValue() + 1)
      failSafe(index)
    }
    if (ByPass.getRange(index, 1).getValue() == true) {
      failSafe(index)
      console.log("bypass attempt " + index)
      runner(index)
      lastIndex = index
    }
  }
  if (lastIndex != 0 && lastIndex < 7) {
    ByPass.getRange(lastIndex, 2).setValue(0)
    ByPass.getRange(lastIndex + 1, 1).setValue("true")
  }
  if (lastIndex != 0 && lastIndex == 7 && ByPass.getRange(lastIndex, 2).getValue() >= 1) {
    ByPass.getRange(lastIndex, 2).setValue(0)
    ByPass.getRange("B1:B5").setValue(0)
  }
   } finally {
    lock.releaseLock();
  }
}

function runner(n) {
  console.log("runner called with n:", n)
  ByPass.getRange(n, 1).setValue("")
  console.log("running API UPDATE " + n)
  if (n == 1) { findNewLevels(); PlayerClearsLocal.sort(2) } //step 1
  if (n == 2) { deletedLevelsHandler(); requestAllListUploaders() } //step 2
  if (n == 3) { updatePlacementsRange("Range", 0, 24) }
if (n == 4) { updatePlacementsRange("Range", 24, 47) }
if (n == 5) { updatePlacementsRange("Range", 43, 70) }
if (n == 6) { 
  var validCount = getAllValidCourseIDs().placements.length;
  updatePlacementsRange("Range", 70, validCount); 
}
if (n == 7) {
  ByPass.getRange(1, 2, 7, 1).setValue(0);
  deleteTrigger();

  var existing = getProjectTriggersByName('createListUpdateTrigger');
  for (var i = 0; i < existing.length; i++) {
    ScriptApp.deleteTrigger(existing[i]);
  }

  // Clean up old leaderboardUpdatev2 triggers before adding new one
  var lbTriggers = getProjectTriggersByName('leaderboardUpdatev2');
  for (var i = 0; i < lbTriggers.length; i++) {
    ScriptApp.deleteTrigger(lbTriggers[i]);
  }

  ScriptApp.newTrigger("leaderboardUpdatev2")
    .timeBased()
    .after(60000)
    .create();

  ScriptApp.newTrigger("createListUpdateTrigger")
    .timeBased()
    .everyHours(12)
    .create();
}
}

function deleteTrigger() {
  var triggers = getProjectTriggersByName('updateListEveryMin');
  for (var i = 0; i < triggers.length; ++i) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function getProjectTriggersByName(name) {
  return ScriptApp.getProjectTriggers().filter(
    function (s) {
      return s.getHandlerFunction() === name;
    });
}

function debugLeaderboardUpdate() {
  console.log("=== DIAGNOSING LEADERBOARD UPDATE ===");
  
  // Get the PID of the player who should have updated points
  var pidToCheck = 120666286361882; // Replace with your player's PID
  
  // Check if the player has cleared the level in PlayerClearsLocal
  var finder = PlayerClearsLocal.createTextFinder(pidToCheck.toString());
  var found = finder.findAll();
  console.log("Player appears in PlayerClearsLocal:", found.length, "times");
  
  // Check if the player's points were actually recalculated
  var pidTable = PIDTable.getRange(1, 1, PIDTable.getLastRow(), 3).getValues();
  for (var i = 0; i < pidTable.length; i++) {
    if (pidTable[i][0] == pidToCheck) {
      console.log("Player found in PID Table at row", i+1);
      console.log("Current points:", pidTable[i][2]);
      break;
    }
  }
  
  // Run the full update
  console.log("Running leaderboardUpdatev2()...");
  leaderboardUpdatev2();
  
  console.log("=== UPDATE COMPLETE. Check the leaderboard now. ===");
}

function setupWatchdogTrigger() {
  // Run this ONCE manually to install a watchdog that runs every hour
  var existing = getProjectTriggersByName('restart');
  if (existing.length === 0) {
    ScriptApp.newTrigger("restart")
      .timeBased()
      .everyHours(1)
      .create();
  }
}

function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function debugfunc(){
  runner(1)
  

  
  
  
}

function auditPlacementColumn() {
  var data = PlayerClearsLocal.getRange(1, 2, PlayerClearsLocal.getLastRow()).getValues();
  for (var i = 0; i < data.length; i++) {
    console.log("Row " + (i+1) + ": " + data[i][0]);
  }
}


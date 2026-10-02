function pidTableFillerV2() {
  //create a list of all PIDs on the PID sheet
  var pidList = PIDTable.getRange(1, 1, PIDTable.getLastRow()).getValues().map(function(row) {return row[0];})
  //invalidPIDList.forEach(i => pidList.removeByValue(i))
  console.log(pidList) 


  //parse every level for new PIDs
  var rowCount = pidList.length + 1
  for (var levelIndex = 1; levelIndex <= 100; levelIndex++) {
    var index = 3
    while (!PlayerClearsLocal.getRange(levelIndex,index).isBlank()) {

      //extract the PID and check if its a duplicate
      //console.log(extractThePID(PlayerClearsLocal.getRange(levelIndex,index).getValue()))
      var playerID = extractThePID(PlayerClearsLocal.getRange(levelIndex,index).getValue())
      if (invalidPIDList.includes(parseInt(playerID))) { 
        //if the account is on the black list
        console.log("black listed PID found: " + playerID + " skipping this PID")
      } else if (playerID == "") {  
        console.log("empty PID somehow so we skip")
      } else if (!pidList.includes(parseInt(playerID))) {
        //PID not on the PID list found
        console.log("new pid found: " + playerID)
        pidList.push(playerID)
        var playerRecordCells = PlayerClearsLocal.createTextFinder(playerID).findAll().map(i => ({row:i.getRow(), col:i.getColumn()}))
        var correctionRecord = "XX Null             d"
        //create a correction record to fill in for their name
        for (var i = 0; i < playerRecordCells.length; i++) {
          var cellContents = PlayerClearsLocal.getRange(playerRecordCells[i].row, playerRecordCells[i].col).getValue()
          if (!cellContents.includes("Null")) {
            correctionRecord = cellContents;
            break;
          }
        }
  
        var pidPos = correctionRecord.search(playerID)
        //print the correction record and their PID on the table
        PIDTable.getRange(rowCount,1).setValue(playerID)
        PIDTable.getRange(rowCount,2).setValue(correctionRecord.slice(0, pidPos))
        rowCount++
      }
      index++
    }
  }
}

function exFuncV2(exRow) {
  if (exRow != null) {
    try {
      var exOut = PlayerClearsLocal.createTextFinder(Exceptions.getRange(exRow, 1).getValue()).findNext().getRow()
      return exOut 
    } catch(err) {
      console.error(err)
      console.log(Exceptions.getRange(exRow, 1).getValue() + " is not on the list anymore")
      return null
    }
  }
}

function findAllIndexesOfValue(arr2d, pid) {
  const indexes = []; 
  const placements = [];

  for (let rowIndex = 0; rowIndex < arr2d.length; rowIndex++) {
    for (let colIndex = 0; colIndex < arr2d[rowIndex].length; colIndex++) {
      if (arr2d[rowIndex][colIndex].includes(pid)) {
        //console.log(rowIndex)
        indexes.push([rowIndex,colIndex]);
        placements.push(rowIndex + 1)
      }
    }
  }
  return [placements, indexes];
}

function listArrayCleanUp(listArray, indexes) {
  for (i=0; i < indexes.length; i++) {
    //console.log(indexes[i][0] + " " + indexes[i][1] + " removed")
    listArray[indexes[i][0]].splice(indexes[i][1],1)
  }
}

function pidLeaderboardUpdateV4() {
  var allPoints = []
  const playerClearsLocalStorage = PlayerClearsLocal.getRange(1, 3, 100, PlayerClearsLocal.getLastColumn()).getValues()
  const pidsArray = PIDTable.getRange(1, 1, PIDTable.getLastRow()).getValues().map(i => i[0])

  for (var pidIndex = 0; pidIndex < pidsArray.length; pidIndex++) {
    var clears = []
    for (var row = 0; row < 100; row++) {
      var col = 0
      while (playerClearsLocalStorage[row][col] != '') {
        if (playerClearsLocalStorage[row][col].includes(pidsArray[pidIndex] + '')) {
          clears.push(row + 1)
        }
        col++
      }
    }
    var points = 0
    for (var clearsIndex = 0; clearsIndex < clears.length; clearsIndex++) {
      points += pointFunction(clears[clearsIndex])
    }
    allPoints.push([points])
  }

  console.log(allPoints)
  PIDTable.getRange(1, 3, PIDTable.getLastRow()).setValues(allPoints)
  run()
}

function leaderboardUpdatev2() {
  pidTableFillerV2()
  //pidLeaderboardUpdateV2()
  pidLeaderboardUpdateV4()
}


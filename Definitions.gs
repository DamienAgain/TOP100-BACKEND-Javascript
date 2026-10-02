var App = SpreadsheetApp //this is just kinda needed for any kind of App Script aplication, it sets up our spreadsheet in the directory of my google drive
var SS = App.getActiveSpreadsheet()  //SS is a set up variable that calls for the spread sheet so we can do things with it in later lines
var ListPage = SS.getSheetByName("List") //this is the page where the list is setup 
//var PlayerClears = SS.getSheetByName("Player Clears")
var Leaderboard = SS.getSheetByName("Leaderboard")
//var PlayersListing = SS.getSheetByName("Players Listing")
var PIDTable = SS.getSheetByName("PID Conversion Table")
var Exceptions = SS.getSheetByName("Exceptions")
var Form = SS.getSheetByName("Form Responses")
var ByPass = SS.getSheetByName("Index")
var PlayerClearsLocal = SS.getSheetByName("Player Clears Local Storage")
var DeletedLevelData = SS.getSheetByName("DeletedLevelData")
var TestPage = SS.getSheetByName("Test Page")
var ChangeLog = SS.getSheetByName("Change Log")
var ButEasySheet = SS.getSheetByName("ButEasyHandler")
var TestPIDTable = SS.getSheetByName("testPIDTable")

var date = new Date();
var formatedDate = (date.getDate() + " " + Utilities.formatDate(date, 'GMT', 'MMMM') + " " + date.getFullYear())

//list of all alt accounts and banned PIDs
var  invalidPIDList = [181923774526828, 141366038381571, 131994173632623, 83368521707129, 166543127816074, 47311779342184, 34013311733390, 154478805304085, 120666286361882, 11277702609971, 44415782311305, 167855126211290, 54502172264151];
//US Booliam 120666286361882
Array.prototype.removeByValue = function (val) {
  for (var i = 0; i < this.length; i++) {
    if (this[i] === val) {
      this.splice(i, 1);
      i--;
    }
  }
  return this;
}

function pointFunction(placement) {
  if (1 <= placement && placement <= 5) {
    return (-125 * placement) + 1125
  } else if (5 < placement && placement <= 20) {
    return (538*((0.9642)**(placement-1)))
  }
  else if (20 < placement && placement <= 50) {
    return (396.0208*((Math.E)**(-0.023*placement)))
  }
  else if (placement > 50) {
    return (423.5*((Math.E)**(-0.02493*placement))+3)
  } 
  else { 
    return "bruh"
  }
}

/* Old Point System, Jul/6/2025
function pointFunction(placement) {
  if (placement <= 20) {
    return (500*((0.9642)**(placement-1)))
  }
  else if (20 < placement && placement <= 50) {
    return (396.0208*((Math.E)**(-0.023*placement)))
  }
  else if (placement > 50) {
    return (1523*((Math.E)**(-0.05*placement)))
  } 
  else { 
    return "input "+ placement + " is out of range"
  }
}
*/
function run() {
  for (var i=1; i<101; i++) {
    //console.log(Math.floor(points))
    ListPage.getRange(i,5).setValue(pointFunction(i))
    //console.log(ListPage.getRange(i,5).getValues());
  }
  leaderboardColorCoder()
  PlayerClearsLocal.sort(2)
}

function leaderboardColorCoder() {
  goatBool = false
  legendBool = false
  grandMasterBool = false
  masterBool = false
  profBool = false
  advancedBool = false
  interBool = false
  beginnerBool = false
  var playerPoints = Leaderboard.getRange("B1:B").getValues()
  Leaderboard.getRange("C1:C").setValue('') //clear all of col C

  for (var i = 0; i < playerPoints.length; i++) {
    if (playerPoints[i] >= 5000) {
      Leaderboard.getRange(i+1,1,1,3).setBackground("magenta") 
      goatBool = true
    } else if (playerPoints[i] >= 2500) {
      if (goatBool && i > 0) {
        Leaderboard.getRange(i, 3).setValue("Goat")
        goatBool = false
      }
      Leaderboard.getRange(i+1,1,1,3).setBackground("cyan")
      legendBool = true
    } else if (playerPoints[i] >= 1000) {
      if (legendBool && i > 0) {
        Leaderboard.getRange(i, 3).setValue("Legend")
        legendBool = false
      }
      Leaderboard.getRange(i+1,1,1,3).setBackground("#8e7cc3")
      grandMasterBool = true
    } else if (playerPoints[i] >= 500) {
      if (grandMasterBool && i > 0) {
        Leaderboard.getRange(i, 3).setValue("Grand Master")
        grandMasterBool = false
      }
      Leaderboard.getRange(i+1,1,1,3).setBackground("#93c47d")
      masterBool = true
    } else if (playerPoints[i] >= 250) {
      if (masterBool && i > 0) {
        Leaderboard.getRange(i, 3).setValue("Master")
        masterBool = false
      }
      Leaderboard.getRange(i+1,1,1,3).setBackground("#e06666")
      profBool = true
    } else if (playerPoints[i] >= 100) {
      if (profBool && i > 0) {
        Leaderboard.getRange(i, 3).setValue("Professional")
        profBool = false
      }
      Leaderboard.getRange(i+1,1,1,3).setBackground("#ff9900")
      advancedBool = true
    } else if (playerPoints[i] >= 50) {
      if (advancedBool && i > 0) {
        Leaderboard.getRange(i, 3).setValue("Advanced")
        advancedBool = false
      }
      Leaderboard.getRange(i+1,1,1,3).setBackground("yellow")
      interBool = true
    } else if (playerPoints[i] >= 10) {
      if (interBool && i > 0) {
        Leaderboard.getRange(i, 3).setValue("Intermediate")
        interBool = false
      }
      Leaderboard.getRange(i+1,1,1,3).setBackground("#00ff00")
      beginnerBool = true
    } else if (playerPoints[i] >= 0) {
      if (beginnerBool && i > 0) {
        Leaderboard.getRange(i, 3).setValue("Novice")
        beginnerBool = false
      }
      Leaderboard.getRange(i+1,1,1,3).setBackground("white")
    }
  }
}

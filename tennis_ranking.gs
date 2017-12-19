// Available tournament categories
var categories = [] ; 

// tournament 1st round starts at column C , 0 in int
var startColumn = 0;
var finalColumn = 0;

// tournament players start after 4rth row
var startingRow=3;

// points awarded in each round
var columnsRoundPoints = [ ];

// A list of new created docs;
var newDocs = [] ;
// get all categories from categories sheet(2nd) 
// and store them in categories array
function getAllCategories() {
   
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(ss.getSheets()[1]);
  var range = ss.getDataRange();
  var values = range.getValues();
  var cat_name="";
  
 for( var i=0; i < values.length ; i++ ) {
    name = values[i][0]; // row/column
    categories.push(name);
  }
  
  SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Ranking')
      .addItem('Calculate All', 'menuItem1')
      .addSeparator()
      .addSubMenu(ui.createMenu('Category')
          .addItem('Calculate for a specific category', 'menuItem2'))
      .addToUi();
}

function menuItem1() {
//  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
//     .alert('You clicked the first menu item!');
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Επιλέξατε υπολογισμό βαθμολογίας όλων των κατηγοριών',
     'Επιβεβαιώστε για συνέχεια',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    //ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    return;
  }
  
  // get all available categories
  getAllCategories();

  //SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert("categories.length="+categories.length+ " ");
     
  var i=0;  
  for( i=0 ; i< categories.length; i++ ) {
    ////////////////////////Logger.log("category="+categories[i]);
    //SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    // .alert("cat = "+categories[i]);
    
    calculateCategoryRanking( categories[i] );
  
  }
 
 var editors = SpreadsheetApp.getActiveSpreadsheet().getEditors();
 
  var emailDescription="Πραγματοποιήθηκε αυτόματη παραγωγή συγκεντρωτικής βαθμολογίας για τις παρακάτω κατηγορίες \n\n";
  for( i=0 ; i < newDocs.length; i++ ){
    var newdoc = newDocs[i];
    // create the email description
    emailDescription+="κατηγορία: "+newdoc.category+" όνομα αρχείου: "+newdoc.name+" URL: "+newdoc.url+"\n\n";
    
    //share new doc with competition_tournament sheet editors
    var new_ss = SpreadsheetApp.openByUrl(newdoc.url);
     for(var j=0; j< editors.length ; j++) {
       new_ss.addEditor(editors[j].getEmail()) ;
     }
     
  }
  MailApp.sendEmail(SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(), "αυτόματη παραγωγή συγκεντρωτικής βαθμολογίας", emailDescription);
  
  
 for( i = 0 ; i< editors.length; i++ ) {
   MailApp.sendEmail(editors[i].getEmail(), "αυτόματη παραγωγή συγκεντρωτικής βαθμολογίας", emailDescription);
 }
  
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('ΟΙ ΝΕΕΣ ΣΥΓΚΕΝΤΡΩΤΙΚΕΣ ΒΑΘΜΟΛΟΓΙΕΣ ΥΠΟΛΟΓΙΣΤΗΚΑΝ ΜΕ ΕΠΙΤΥΧΙΑ! ΠΑΡΑΚΑΛΩ ΕΛΕΓΞΤΕ ΤΟ EMAIL ΣΑΣ ΓΙΑ ΛΕΠΤΟΜΕΡΕΙΕΣ');

}

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the second menu item!');
}

// scan all sheet cells until finding word "Seed" or "seed", Then set starting row and columns accordingly(next row and next column)
function getSheetTournamentBoundaries(values) {

  for( var i=0; i < values.length ; i++ ) {
    for (var j = 0; j < values[i].length; j++){
      if ( ( values[i][j] != "" ) &&
        ( values[i][j] == "seed" || values[i][j] == "Seed" ) )
      {
        startColumn = j+1;
        startingRow = i+1;
         return;
      }
    }
  }
  
}

// Calculates the total ranking of a category
function calculateCategoryRanking( category ) {

  // scan all rows and get all documents for this category
  // A:tournament, B:category, C:date, D:grade , E:SHEET
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var compSheet = ss.getSheetByName("competitions");
  SpreadsheetApp.setActiveSheet(compSheet);
  var range = ss.getDataRange();
  var values = range.getValues();
  var cat_name="";
  
  var tournament="";
  var category_tourn="";
  var date_tourn="";
  var grade_tourn="";
  var sheet_tourn="";
  
  var total_players_ranking = [];
  
  for( var i=1; i < values.length ; i++ ) {
    // get a row
    tournament=values[i][0];
    category_tourn=values[i][1];
    date_tourn=values[i][2];
    grade_tourn=values[i][3];
    sheet_tourn=values[i][4];
    type_tourn=values[i][5];
    
    type_tourn = getSingleOrDouble(category);
    
    var players_ranking = [];
    //Logger.log( "tournament="+tournament+" category_tourn="+category_tourn+" date_tourn="+date_tourn+" grade_tourn="+grade_tourn+" sheet_tourn="+sheet_tourn );
    if ( category == category_tourn )  {
      
      if( tournament == "BASE" ) 
      {
        calculateBaseRanking(sheet_tourn, players_ranking ) ;
      }
      else
        calculateTournamentRanking( grade_tourn, sheet_tourn, type_tourn, players_ranking );
      
      // log the ranking of the tournament
      //Logger.log( "Performed ranking for tournament "+tournament );
      
      // Update the categry ranking
      for( var j =0 ; j< players_ranking.length ; j++) {
        Logger.log("player="+players_ranking[j].name+" points="+players_ranking[j].points);
        updateTotalRanking(players_ranking[j] , total_players_ranking);
      }
    }
    
      
  }
  
  if( total_players_ranking.length != 0 ) {
    createNewSheet( category , total_players_ranking );
  }

}

function createNewSheet( category_tourn, total_players_ranking) {
  for( var j =0 ; j< total_players_ranking.length ; j++) {
        //Logger.log("*** TOTAL player="+total_players_ranking[j].name+" points="+total_players_ranking[j].points);
       
  }
  
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth()+1; 
  var yyyy = today.getFullYear();
  var min=today.getMinutes();
  var hour=today.getHours();
  var sec=today.getSeconds();

  if(dd<10) {
    dd='0'+dd
  } 

  if(mm<10) {
    mm='0'+mm
  } 

  today = mm+'_'+dd+'_'+yyyy+"_"+hour+":"+min+":"+sec;
  var spreadsheet_name="total_"+category_tourn+"_"+yyyy;
  // Check  if spreadsheet exists already
  var FileIterator = DriveApp.getFilesByName(spreadsheet_name);
  var ssNew;
  if (FileIterator.hasNext())
  {
    var file = FileIterator.next();
    ////Logger.log( "filename="+file.getName()+ "* sheet_tourn="+sheet_tourn+"*" );
       
     ssNew = SpreadsheetApp.open(file);
   }
  // end of check
  else {
    ssNew = SpreadsheetApp.create(spreadsheet_name);
    
  }
  //Logger.log("URL="+ssNew.getUrl()+ " NAME="+spreadsheet_name);
  
  // push doc details in newDoc so we will send an email later
  var newdoc = new Object();
  newdoc.url = ssNew.getUrl();
  newdoc.category=category_tourn;
  newdoc.name=spreadsheet_name;
  
  newDocs.push(newdoc);
  
  // create headers and popuate with data
  //var ss = SpreadsheetApp.openByUrl(ssNew.getUrl());
  var sheet = ssNew.getSheets()[0];
  
  sheet.clear();
  
  var description="ΣΥΓΚΕΝΤΡΩΤΙΚΗ ΒΑΘΜΟΛΟΓΙΑ ΚΑΤΗΓΟΡΙΑΣ "+category_tourn;
  var description_arr = [];
  
  description_arr.push(description);
  description_arr.push(" ");
  sheet.appendRow(description_arr);
  sheet.appendRow(["RANK", "ONOΜΑ", "ΒΑΘΜΟΛΟΓΙΑ"]);
  var rank=0;
  var num_of_ranked=0;
  for( var j =0 ; j< total_players_ranking.length ; j++) {
     if(  total_players_ranking[j].points > 0   ) {
       sheet.appendRow([ rank, total_players_ranking[j].name, total_players_ranking[j].points] );
       num_of_ranked++;
     }
  }
  
  var row_end_rank=num_of_ranked+3;
  // sort by column 2
  var act_sheet = ssNew.getSheets()[0];
  var range = act_sheet.getRange("A3:C"+(row_end_rank) );
  range.sort({column: 3, ascending: false});
  act_sheet.getRange("A1:F1").merge();
  
  // calculate rank
  var rank=0;
  var prev_points=999999;
  var cur_points=0;
  
  var values = act_sheet.getDataRange().getValues();
  for (var j=2 ; j<values.length ; j++ ) {
  
    cur_points=parseInt( values[j][2] );
    //Logger.log("val = "+cur_points);
    if( cur_points != prev_points ) {
      rank++;
      prev_points = cur_points;
    }
    act_sheet.getRange(j+1,1).setValue(""+rank);
  }
  
}

function updateTotalRanking(player, total_players_ranking) {
  var found = false;
  
  // if no win the no points
  if( player.has_win == 0 )
    player.points = 0 ;
    
  for( var i = 0 ; i<total_players_ranking.length ; i++ ) {
    var t_player = total_players_ranking[i];
    if ( t_player.name == player.name )  {
      t_player.points+=player.points;
      Logger.log("updateTotalRanking update "+t_player.name+" POINTS="+t_player.points);
      found = true;
    }
  }
  
  if( found == false ){
    Logger.log("updateTotalRanking push "+player.name+" POINTS="+player.points);
    total_players_ranking.push(player);
  }
  
}


function calculateTournamentRanking( grade_tourn, sheet_tourn, type_tourn ,players_ranking ) {

  //if (x.key !== undefined)

  // Open the file
 // var FileIterator = DriveApp.getFilesByName(sheet_tourn);
 // while (FileIterator.hasNext())
  //{
    //var file = FileIterator.next();
    ////Logger.log( "filename="+file.getName()+ "* sheet_tourn="+sheet_tourn+"*" );
    
    //if (file.getName() == sheet_tourn )
    if( sheet_tourn != "" )
    {
     // var sheet = SpreadsheetApp.open(file);
      var sheet = SpreadsheetApp.openByUrl(sheet_tourn);
      //var sheet = sprsheet.getSheets()[0];
      //var fileID = file.getId();
      
      var dataRange = sheet.getDataRange();
      var data = dataRange.getValues();
      
      getSheetTournamentBoundaries(data);
      // get number of players
            
      var num_of_players = getNumberOfPlayers(data);
      //Logger.log( "num_of_players="+num_of_players );
      if( num_of_players == 0 && (grade_tourn != "GROUPS") ) {
        return;
      }
      
      switch( grade_tourn ) {
        case "A" :
          tournGradeA(num_of_players);
          break;
         case "B" :
           tournGradeB(num_of_players);
           break;
         case "C" :
           tournGradeC(num_of_players);
           break;
         case "GROUPS" :
           calculateGroupsRanking(data, players_ranking);
           return;
         default :
           tournGradeA(num_of_players);
      }
      
      setTournamentRounds(num_of_players);
      
      var funcCall = type_tourn+"_"+num_of_players+"_fn( sheet.getSheets()[0], players_ranking)" 
      eval(funcCall);    
      return;
      // start with the winner
      var currentRound = finalColumn;
      //Logger.log("finalColumn="+finalColumn);
      while( currentRound >= startColumn ) {
        
        // for each player of this round give it the points if not done already
        for ( var i=startingRow; i< data.length ; i++ ) {
          cell = data[i][currentRound];
          var pattern = /[0-9]+/;
          if (pattern.test(cell)){
            continue;
          }
         // //Logger.log("cell="+cell+" i="+i+" currentRound="+currentRound);
          if ( ( cell != "" ) && 
            //( cell != "WO" ) &&
            ( cell != "wo" ) &&
            ( cell != "BYE" ) &&
            ( cell != "ret" ) &&
            ( cell != "RET" ) &&
            ( cell != "ΒΥΕ" ) && 
            ( cell != "W.O" )&& 
            ( cell != "W/O" )&& 
            ( cell != "w/o" ) ) {
            if( checkPlayerIsRanked(cell, players_ranking) == false ) {
              var player = new Object();
              player.name = cell;
              player.points = columnsRoundPoints[currentRound];
              
              players_ranking.push(player);
            }
          }
        }  
      
      
        currentRound--;
      }
      
      
    }
  //}
   // Check if there is a  - or MULTILINE and split names and recalc players+ranking
   singlifyDoublePlayers(players_ranking);
}   


function Single_8_fn( sheet, players_ranking ) {
  // We know that the tournament is between E7 - L21
  var range = sheet.getRange(7, 5, 15, 8);
  var values = range.getValues();
  
  // Columns that contain player data, starting from winner (L which is the no 8 in the range )
  var columns_valid = [ 7 , 5 , 3 , 0 ];
  var cell="";
  var currentRound = 3;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
      cell = values[i][columns_valid[a]];
      ////Logger.log(cell+" row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          
          var player = new Object();
          player.name = cell.trim();
          player.points = columnsRoundPoints[currentRound];
          player.has_win = 0;
          players_ranking.push(player);
      }
      
      if( checkHasScore( values,i,columns_valid[a] ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  matchNames( 3 , values , players_ranking, 1);
  
 }

function Double_8_fn( sheet, players_ranking ) {
  // We know that the tournament is between E7 - K37
  var range = sheet.getRange(7, 5, 30, 7);
  var values = range.getValues();
  
  // Columns that contain player data, starting from winner ( K which is the no 7 in the range )
  var columns_valid = [ 6 , 4 , 2 , 0 ];
  var cell="";
  var currentRound = 3;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
      cell = values[i][columns_valid[a]];
      ////Logger.log(cell+" row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          Logger.log("Double8 "+cell+" row="+i+" col="+columns_valid[a]+" round="+currentRound);
          var player = new Object();
          player.name = cell.trim();
          player.points = columnsRoundPoints[currentRound];
          player.has_win = 0;
          players_ranking.push(player);
      }
      
      if( checkHasScore( values,i,columns_valid[a] ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  
  matchNames( 2 , values , players_ranking, 2);


}


function Single_16_fn( sheet, players_ranking ) {
  // We know that the tournament is between E7 - Ν22
  var range = sheet.getRange(7, 5, 31, 10);
  var values = range.getValues();
  
  // Columns that contain player data, starting from winner (Ν which is the no 10 in the range )
  var columns_valid = [9,  7 , 5 , 3 , 0 ];
  var cell="";
  var currentRound = 4;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
      cell = values[i][columns_valid[a]];
      
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
           Logger.log(cell+" row="+i+" col="+columns_valid[a]+" round="+currentRound);
          var player = new Object();
          player.name = cell.trim();
          player.points = columnsRoundPoints[currentRound];
          player.has_win = 0;
          players_ranking.push(player);
      }
      
      if( checkHasScore( values,i,columns_valid[a] ) ) {
        Logger.log("Mark as having win "+cell+" row="+i+" col="+columns_valid[a]+" round="+currentRound);
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  matchNames( 3 , values , players_ranking, 1);
  
}

function Double_16_fn( sheet, players_ranking ) {

  // We know that the tournament is between E7 - M69
  var range = sheet.getRange(7, 5, 63, 9);
  var values = range.getValues();
  
  // Columns that contain player data, starting from winner (Ν which is the no 10 in the range )
  var columns_valid = [8,  6 , 4 , 2 , 0 ];
  var cell="";
  var currentRound = 4;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
      cell = values[i][columns_valid[a]];
      ////Logger.log(cell+" row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          
          var player = new Object();
          player.name = cell.trim();
          player.points = columnsRoundPoints[currentRound];
          player.has_win = 0 ;
          players_ranking.push(player);
      }
      
      if( checkHasScore( values,i,columns_valid[a] ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  matchNames( 2 , values , players_ranking ,2);
  
}


function Single_32_fn( sheet, players_ranking ) {
  // We know that the tournament is between E7 - Ν22
  var range = sheet.getRange(7, 5, 64, 10);
  var values = range.getValues();
  var cell="";
  
  // In 32 the N column contains the winner
    cell = values[31][9];
  var currentRound = 5;
  
  if( checkValidPlayerEntry(cell) == true ) {
    var player = new Object();
    player.name = cell.trim();
    player.has_win=0;
    player.points = columnsRoundPoints[currentRound];
    players_ranking.push(player);
  }
  if( checkHasScore( values,31, 9 ) ) {
        hasWin(cell, players_ranking);
  }
  // Columns that contain player data, starting from winner (Ν which is the no 10 in the range )
  var columns_valid = [ 9, 7 , 5 , 3 , 0 ];
  
  currentRound = 4;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
      cell = (values[i][columns_valid[a]]).trim();
      ////Logger.log(cell+" row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          
          var player = new Object();
          player.name = cell;
          player.points = columnsRoundPoints[currentRound];
          player.has_win=0;
          players_ranking.push(player);
      }
      
      if( checkHasScore( values,i,columns_valid[a] ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  matchNames( 3 , values , players_ranking, 1);
  
}


function Double_32_fn( sheet, players_ranking ) {
  
  // We know that the 1st part of the tableu of the tournament is between E7 - Ν69
  var range = sheet.getRange(7, 5, 63, 10);
  var values = range.getValues();
  var cell="";
  
  // In 32 the N column contains the winners N66 and N67
  cell = values[59][9].trim();
  var currentRound = 5;
  //Logger.log("Winner N66="+cell);
  if( checkValidPlayerEntry(cell) == true ) {
    var player = new Object();
    player.name = cell.trim();
    player.has_win = 0;
    player.points = columnsRoundPoints[currentRound];
    players_ranking.push(player);
    //Logger.log("Winner N66="+player.name+"points="+player.points);
  }
  if( checkHasScore( values,59,9 ) ) {
    hasWin(cell, players_ranking);
  }
  
  cell = values[60][9].trim();
  currentRound = 5;
  //Logger.log("Winner N67="+cell);
  if( checkValidPlayerEntry(cell) == true ) {
    var player = new Object();
    player.name = cell.trim();
    player.has_win=0;
    player.points = columnsRoundPoints[currentRound];
    players_ranking.push(player);
    //Logger.log("Winner N67="+player.name+"points="+player.points);
  }
  if( checkHasScore( values,60,9 ) ) {
    hasWin(cell, players_ranking);
  }
  // Semi final is between L64 and L69
  
  for(var i=57 ; i<=62 ; i++ ) {
    cell = values[i][7];
    currentRound = 4;
    //Logger.log("Semi name="+i+" "+cell); 
    if ( ( checkValidPlayerEntry(cell) == true ) &&
         ( checkPlayerIsRanked(cell, players_ranking) == false ) )  {
      var player = new Object();
      player.name = cell.trim();
      player.has_win=0;
      player.points = columnsRoundPoints[currentRound];
      players_ranking.push(player);
      //Logger.log("Semi N67="+player.name+"points="+player.points);
    }
    if( checkHasScore( values, i, 7 ) ) {
      hasWin(cell, players_ranking);
    }
  }
  
  // Columns that contain player data
  var columns_valid = [  7 , 5 , 3 , 0 ];
  
  currentRound = 3;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
    
      cell = (values[i][columns_valid[a]]).trim();
      //Logger.log(cell+"Double 32 row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          
          var player = new Object();
          player.name = cell;
          player.points = columnsRoundPoints[currentRound];
          player.has_win=0;
          players_ranking.push(player);
      }
      if( checkHasScore( values, i, columns_valid[a]  ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  matchNames( 3 , values , players_ranking, 2);
  
  // Now get the 2nd part of tableu E82 - N144
  
  range = sheet.getRange(82, 5, 63, 10);
  values = range.getValues();
  cell="";
  
  //var columns_valid = [  7 , 5 , 3 , 0 ];
  
  currentRound = 3;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
    
      cell = (values[i][columns_valid[a]]).trim();
      Logger.log(cell+"Double 32 row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          Logger.log(cell+"Double 32 row="+i+" col="+columns_valid[a]+" round="+currentRound);
          var player = new Object();
          player.name = cell;
          player.points = columnsRoundPoints[currentRound];
          player.has_win=0;
          players_ranking.push(player);
      }
      if( checkHasScore( values, i, columns_valid[a]  ) ) {
        hasWin(cell, players_ranking);
     }
    
    }
    currentRound--;
  }
  matchNames( 3 , values , players_ranking, 2);
  
}


function Single_64_fn( sheet, players_ranking ) {

  // We know that the tournament is between E7 - p70
  var range = sheet.getRange(7, 5, 64, 12);
  var values = range.getValues();
  var cell="";
  
  // In 64 the P column contains the winner
  cell = values[31][11];
  var currentRound = 6;
  
  if( checkValidPlayerEntry(cell) == true ) {
    var player = new Object();
    player.name = cell.trim();
    player.has_win=0;
    player.points = columnsRoundPoints[currentRound];
    players_ranking.push(player);
  }
  
  if( checkHasScore( values, 31, 11  ) ) {
    hasWin(cell, players_ranking);
  }
  // The semi final are in  n37 and n39
  var currentRound = 5;
  
  cell = values[30][9].trim();
  if( checkValidPlayerEntry(cell) == true &&
    checkPlayerIsRanked(cell, players_ranking) == false ) {
  
    var player = new Object();
    player.name = cell.trim();
    player.has_win=0;
    player.points = columnsRoundPoints[currentRound];
    players_ranking.push(player);
  }
  if( checkHasScore( values, 30, 9  ) ) {
    hasWin(cell, players_ranking);
  }
  
  cell = values[32][9].trim();
   
  if( checkValidPlayerEntry(cell) == true ) {
    var player = new Object();
    player.name = cell.trim();
    player.has_win=0;
    player.points = columnsRoundPoints[currentRound];
    players_ranking.push(player);
  }
  if( checkHasScore( values, 32, 9  ) ) {
    hasWin(cell, players_ranking);
  }
  
  // Columns that contain player data, starting from winner (Ν which is the no 10 in the range )
  var columns_valid = [ 11,  9, 7 , 5 ,  0 ];
  
  currentRound = 4;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
      cell = (values[i][columns_valid[a]]).trim();
      //Logger.log(cell+" row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          
          var player = new Object();
          player.name = cell;
          player.has_win=0;
          player.points = columnsRoundPoints[currentRound];
       
          players_ranking.push(player);
      }
      
      if( checkHasScore( values, i, columns_valid[a]  ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  
  //go to 2nd round and match names to 1st round
  matchNames( 5 , values , players_ranking, 1);
  
}

function Double_64_fn( sheet, players_ranking ) {
  
  // We know that the 1st part of the tableu of the tournament is between I309 - M335
  var range = sheet.getRange(309, 9, 37, 5);
  var values = range.getValues();
  var cell="";
  
  var currentRound = 6;
  var col = 4;
  
  while( col >= 0 ) {
  
    for( var i=0 ; i < values.length ; i++ ) {
      cell = values[i][col];
         
      if ( ( checkValidPlayerEntry(cell) == true ) &&
           ( checkPlayerIsRanked(cell, players_ranking) == false ) )  {
        var player = new Object();
        player.name = cell.trim();
        player.has_win=0;
        player.points = columnsRoundPoints[currentRound];
        players_ranking.push(player);
        //Logger.log("Semi N67="+player.name+"points="+player.points);
      }
      if( checkHasScore( values, i, col  ) ) {
        hasWin(cell, players_ranking);
      }
  
    }
    
    col-=2;
    currentRound--;
  }
  
  // Columns that contain player data
  var columns_valid = [  6 , 4 , 2 , 0 ];
  
  // Get 1st page tableu
  range = sheet.getRange(7, 5, 63, 7);
  values = range.getValues();
  cell="";
  currentRound = 3;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
    
      cell = (values[i][columns_valid[a]]).trim();
      //Logger.log(cell+"Double 32 row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          
          var player = new Object();
          player.name = cell;
          player.points = columnsRoundPoints[currentRound];
          player.has_win=0;
          players_ranking.push(player);
      }
      if( checkHasScore( values, i, columns_valid[a]  ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  
  //go to 2nd round and match names to 1st round
  matchNames( 2 , values , players_ranking ,2 );
  
  
  // Now get the 2nd part of tableu E82 - N144
  range = sheet.getRange(82, 5, 63, 7);
  values = range.getValues();
  cell="";
  currentRound = 3;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
    
      cell = (values[i][columns_valid[a]]).trim();
      //Logger.log(cell+"Double 32 row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          
          var player = new Object();
          player.name = cell;
          player.points = columnsRoundPoints[currentRound];
          player.has_win=0;
          players_ranking.push(player);
      }
      if( checkHasScore( values, i, columns_valid[a]  ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  matchNames( 2 , values , players_ranking ,2);
  
  // Now get the 3rd part of tableu e157
  range = sheet.getRange(157, 5, 63, 7);
  values = range.getValues();
  cell="";
  currentRound = 3;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
    
      cell = (values[i][columns_valid[a]]).trim();
      //Logger.log(cell+"Double 32 row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          
          var player = new Object();
          player.name = cell;
          player.points = columnsRoundPoints[currentRound];
          player.has_win=0;
          players_ranking.push(player);
      }
      if( checkHasScore( values, i, columns_valid[a]  ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  
  matchNames( 2 , values , players_ranking ,2);
  
  // Now get the 4th part of tableu e232
  range = sheet.getRange(232, 5, 63, 7);
  values = range.getValues();
  cell="";
  currentRound = 3;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
    
      cell = (values[i][columns_valid[a]]).trim();
      //Logger.log(cell+"Double 32 row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          
          var player = new Object();
          player.name = cell;
          player.points = columnsRoundPoints[currentRound];
          player.has_win=0;
          players_ranking.push(player);
      }
      if( checkHasScore( values, i, columns_valid[a]  ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  
  matchNames( 2 , values , players_ranking ,2);
}


function Single_128_fn( sheet, players_ranking ) {
  
  // We know that the 1st part of the tableu of the tournament is between e178 - l192
  //var range = sheet.getRange(160, 10, 30, 7);
  var range = sheet.getRange(178, 5, 30, 8);
  var values = range.getValues();
  var cell="";
  
  var currentRound = 7;
  var col = 7;
  // Columns that contain player data
  var columns_valid_8 = [  7, 5 , 3 ,  0 ];
  //while( col >= 0 ) {
  for( var col = 0 ; col<columns_valid_8.length; col++ ) {
    for( var i=0 ; i < values.length ; i++ ) {
      cell = values[i][columns_valid_8[col]];
         
      if ( ( checkValidPlayerEntry(cell) == true ) &&
           ( checkPlayerIsRanked(cell, players_ranking) == false ) )  {
        var player = new Object();
        player.name = cell.trim();
        player.has_win=0;
        player.points = columnsRoundPoints[currentRound];
        players_ranking.push(player);
        //Logger.log("Semi N67="+player.name+"points="+player.points);
      }
      if( checkHasScore( values, i, col  ) ) {
        hasWin(cell, players_ranking);
      }
  
    }
    
    
   
    currentRound--;
  }
  
  // Columns that contain player data
 var columns_valid = [ 9, 7 , 5 , 3 , 0 ];
  
  // Get 1st page tableu
  range = sheet.getRange(7, 5, 64, 10);
  values = range.getValues();
  cell="";
  currentRound = 4;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
    
      cell = (values[i][columns_valid[a]]).trim();
      //Logger.log(cell+"Double 32 row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          
          var player = new Object();
          player.name = cell;
          player.has_win=0;
          player.points = columnsRoundPoints[currentRound];
       
          players_ranking.push(player);
      }
      if( checkHasScore( values, i, columns_valid[a]  ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  matchNames( 3 , values , players_ranking , 1);
  
  // Now get the 2nd part of tableu E84 - N144
  range = sheet.getRange(84, 5, 64, 10);
  values = range.getValues();
  cell="";
  currentRound = 4;
  
  for( var a = 0 ; a<columns_valid.length; a++ ) {
    for( var i = 0 ; i<values.length; i++ ) {
    
      cell = (values[i][columns_valid[a]]).trim();
      //Logger.log(cell+"Double 32 row="+i+" col="+columns_valid[a]+" round="+currentRound);
      if( checkValidPlayerEntry(cell) == true &&
          checkPlayerIsRanked(cell, players_ranking) == false ) {
          
          var player = new Object();
          player.name = cell;
          player.points = columnsRoundPoints[currentRound];
          player.has_win=0;
          players_ranking.push(player);
      }
      if( checkHasScore( values, i, columns_valid[a]  ) ) {
        hasWin(cell, players_ranking);
      }
    }
    currentRound--;
  }
  matchNames( 3 , values , players_ranking);
}


function getSingleOrDouble(category) {

  switch (category) {
  
    case "R- M OPEN":
    case "R- M1 14-29":
    case "R- M2 30-39":
    case "R- M3 40-49":
    case "R- M4 50-59":
    case "R- M5 60+":
    case "R-W-44":
    case "R- W 45+":
    case "ROOKIES":
    case "W1":
    case "W2":
      return "Single";
      break;
    case "R-MD OPEN":
    case "R-MIXED OPEN":
    case "R-MD 14-39":
    case "R-MD 40-54":
    case "R- MD 55+":
    case "WD1":
    case "WD2":
    case "MIXED1":
    case "MIXED 2":
    case "ADULT-CHILD":
      return "Double";
    default:
      return "Double";
  }
}

function checkHasScore(data,row,col) {
  
  var pattern = /[0-9]+/;
  var cell1_down = "";
  var cell2_down = "";
  
  if( (row+1) >= data.length )
    return false;
  cell1_down = data[row+1][col];
  if( pattern.test(cell1_down) || ( cell1_down == "wo" ) || ( cell2_down == "WO" ) || ( cell2_down == "W/O" )|| ( cell2_down == "RET" ) || ( cell2_down == "ret" )  ) {
      return true;
  }
  
  // If the player is marked with WO in the score area
  // it means that he did nog give any game, so no score
  if( (row+2) >= data.length )
    return false;
    
  cell2_down = data[row+2][col];
  if( pattern.test(cell2_down) || ( cell2_down == "wo" ) || ( cell2_down == "WO" ) || ( cell2_down == "W/O" )|| ( cell2_down == "RET" ) || ( cell2_down == "ret" ) ) {
      return true;
  }
    
  return false;
}

function checkValidPlayerEntry(cell) {

  var pattern = /[0-9]+/;
  if (pattern.test(cell)){
    return false;
  }
 //         //Logger.log("cell="+cell+" i="+i+" currentRound="+currentRound);
 if ( ( cell != "" ) && 
   ( cell != "WO" ) &&
   ( cell != "wo" ) &&
   ( cell != "BYE" ) &&
   ( cell != "ret" ) &&
   ( cell != "RET" ) &&
   ( cell != "ΒΥΕ" ) && 
   ( cell != "W.O" )&& 
   ( cell != "W/O" )&& 
   ( cell != "w/o" ) ) {
     return true;
   } else {
     return false;
   }


}

function calculateGroupsRanking( data, players_ranking ) {

  var players_cells = [16, 15, 11, 10 , 9 , 5, 4, 3];
  
  for ( var i = 0; i<players_cells.length; i++ ) {
    
    var name = data[ players_cells[i] ][ 2 ];
    var  points = data[ players_cells[i] ][ 6 ];
    
    if( ( name != "" ) && (name != "X") && (name != "χ" ) && (name != "Χ" ) && (name != "x" ) )
    {
      if( ( points != "" ) && (points != "X") && (points != "χ" ) && (points != "Χ" ) && (points != "x" ) ) {
        //Logger.log("groups - player="+name+" points="+points);
      } else
      {
        continue;
      }
    
    } else {
      continue;
    }
    if( checkPlayerIsRanked(name, players_ranking) == false ) {
      var player = new Object();
      player.name = name;
      player.points = points;
        
      players_ranking.push(player);
    }
  }
  // Check if there is a  - or MULTILINE and split names and recalc players+ranking
   singlifyDoublePlayers(players_ranking);

}
function calculateBaseRanking(sheet_tourn, players_ranking ) {

  // Open the file
 // var FileIterator = DriveApp.getFilesByName(sheet_tourn);
//  while (FileIterator.hasNext())
//  {
  //  var file = FileIterator.next();
 //   //Logger.log( "filename="+file.getName()+ "* sheet_tourn="+sheet_tourn+"*" );
    
   // if (file.getName() == sheet_tourn )
   if( sheet_tourn != "" )
    {
      //var sheet = SpreadsheetApp.open(file);
      var sheet = SpreadsheetApp.openByUrl(sheet_tourn);
      //var sheet = sprsheet.getSheets()[0];
      //var fileID = file.getId();
      
      var dataRange = sheet.getDataRange();
      var data = dataRange.getValues();
      
      for ( var i=2; i< data.length; i++ ) {
        if ( data[i][1] != "" ) {
          var player = new Object();
          player.name = data[i][1];
          player.points=parseInt(data[i][2]);
          players_ranking.push(player);
        }
      }
      
    }
    
  // }
}

function singlifyDoublePlayers(players_ranking) {
  var partner_players_ranking = [];
  // This how double players teammates can be split
  var delimiters=['-', '\n' ,'\\','/' ];
  
  for ( var i=0; i< players_ranking.length ; i++ ) {
    var player = players_ranking[i];
    for(var j = 0 ;j< delimiters.length ; j++ ) {
    
      if ( player.name.indexOf(delimiters[j]) != -1 ) {
      
        var names = player.name.split(delimiters[j]);
    
        player.name = names[0].trim();
        var partner=new Object();
        partner.name = names[1].trim();
        partner.points = player.points;
      
        partner_players_ranking.push(partner);
      }
    }
  
  }
  
  for ( var j=0; j < partner_players_ranking.length ; j++ ) {
    var partner = partner_players_ranking[j];
    players_ranking.push(partner);
  }
  
}

function checkPlayerIsRanked(cell, players_ranking) {

  for( var i = 0 ; i< players_ranking.length ; i++ ) {
    var player = players_ranking[i];
    
    if( player.name == cell ) {
      return true;
    }
  
  }
  
  return false;
}

function hasWin(cell, players_ranking) {

  for( var i = 0 ; i< players_ranking.length ; i++ ) {
    var player = players_ranking[i];
    
    if( player.name == cell ) {
      player.has_win = 1;
    }
  
  }
  
}

function matchNames( round2 , values , players_ranking , player_per_team) {

  var cell="";
  var fullName="";
  // Scan all round2 names
  for( var i = 0 ; i<values.length ; i++ ) {
  
    cell = values[i][round2];
    if( checkValidPlayerEntry(cell) ){
        fullName=getRound1Name(values, i,  cell, player_per_team );
        if ( (fullName!="") && (fullName != cell ) ){
          changeName(cell,  fullName, players_ranking);
        }
    }
    
  }

}


function getRound1Name(values,row, value, player_per_team ) {

  //get current name in 2nd round
  var cell = "";
  var match_name="";
  
  // go up until meeting num_of_names or reach 0
  var num_of_players=0;
  for( var i = row ; i>=0 ; i-- ) {
    cell = values[i][0].trim();
    if( checkValidPlayerEntry(cell) ){
      //check if value is a substring of current value
      //Logger.log("getRound1Name: cell="+cell+ "value="+value);
      if( cell.indexOf(value) !=-1 )
        match_name = cell;
        
    }
    
    if( cell != "" )
      num_of_players++;
      
    if( player_per_team == num_of_players )
      break;
  }
  
  num_of_players=0;
  // go down until meeting num_of_names
  for( var i = row+1 ; i<values.length ; i++ ) {
    cell = values[i][0].trim();
    if( checkValidPlayerEntry(cell) ){
      //check if value is a substring of current value
       //Logger.log("getRound1Name: cell="+cell+ "value="+value);
      if( cell.indexOf(value) !=-1 )
        match_name = cell;
    }
    
    if( cell != "" )
      num_of_players++;
      
    if( player_per_team == num_of_players )
      break;
  }
  
  return match_name;
}

function changeName(cur_name, new_name, players_ranking) {

  for( var i = 0 ; i< players_ranking.length ; i++ ) {
    var player = players_ranking[i];
    
    if( player.name == cur_name ) {
      player.name = new_name
    }
  
  }

}

function getNumberOfPlayers(cells) {

  //TODO : Parse cell J2 : MAIN DRAW(64)
  var pattern = /[0-9]+/;
  var num_of_players = 0;
  //Logger.log("getNumberOfPlayeres max lenght="+cells.length);
  // count the names in column A
  for ( var i=startingRow; i< cells.length ; i++ ) {
    cell = cells[i][0];
    //Logger.log("getNumberOfPlayers cell="+cell);
    if ( ( cell != "" ) && ( pattern.test(cell) )&&( cell != "WO" ) && ( cell != "BYE" ) ) {
      num_of_players++;
    }
  }

  Logger.log("getNumberOfPlayers="+num_of_players);
  return num_of_players;

}

// this function sets the lat column of the tournament given the number of players and the starting column

function setTournamentRounds(numOfPlayers) {
   
  if ( numOfPlayers <= 8 ) {
    finalColumn=3+startColumn;
  } else if ( numOfPlayers <= 16 ) {
    finalColumn=4+startColumn;;
  } else if ( numOfPlayers <= 32 ) {
    finalColumn=5+startColumn;;
  } else if ( numOfPlayers <= 64 ) {
    finalColumn=6+startColumn;;
  } else if ( numOfPlayers <= 128 ) {
    finalColumn=7+startColumn;;
  }
    
}



function tournGradeA(numOfPlayers)
{
  // A column is 1st round, so no points
  columnsRoundPoints[0 ] = 0 ;
  // B , C , D
  if ( numOfPlayers <= 8 ) {
    columnsRoundPoints[1] = 30; 
    columnsRoundPoints[2] = 40;
    columnsRoundPoints[3] = 50;
  } 
  // B, C, D , E
  else if ( numOfPlayers <= 16 ) {
    columnsRoundPoints[1] = 30; 
    columnsRoundPoints[2] = 40;
    columnsRoundPoints[3] = 50;
    columnsRoundPoints[4] = 60;
  }
  // B, C, D, E, F
  else if ( numOfPlayers <=32 ) {
    columnsRoundPoints[1] = 30; 
    columnsRoundPoints[2] = 40;
    columnsRoundPoints[3] = 50;
    columnsRoundPoints[4] = 60;
    columnsRoundPoints[5] = 70;
  }  
  // B, C, D, E, F, G
  else if ( numOfPlayers <= 64 ) {
    
    columnsRoundPoints[1] = 30; 
    columnsRoundPoints[2] = 40;
    columnsRoundPoints[3] = 50;
    columnsRoundPoints[4] = 60;
    columnsRoundPoints[5] = 70;
    columnsRoundPoints[6] = 80;
    
  }
  else if ( numOfPlayers <= 128 ) {
    
    columnsRoundPoints[1] = 30; 
    columnsRoundPoints[2] = 40;
    columnsRoundPoints[3] = 50;
    columnsRoundPoints[4] = 60;
    columnsRoundPoints[5] = 70;
    columnsRoundPoints[6] = 80;
    columnsRoundPoints[7] = 90;
    
  }
}

function tournGradeB(numOfPlayers)
{
  // A column is 1st round, so no points
  columnsRoundPoints[0] = 0 ;
  // B , C , D
  if ( numOfPlayers <= 8 ) {
    columnsRoundPoints[1] = 20; 
    columnsRoundPoints[2] = 30;
    columnsRoundPoints[3] = 40;
  } 
  // B, C, D , E
  else if ( numOfPlayers <= 16 ) {
    columnsRoundPoints[1] = 20; 
    columnsRoundPoints[2] = 30;
    columnsRoundPoints[3] = 40;
    columnsRoundPoints[4] = 50;
  }
  // B, C, D, E, F
  else if ( numOfPlayers <=32 ) {
    columnsRoundPoints[1] = 20; 
    columnsRoundPoints[2] = 30;
    columnsRoundPoints[3] = 40;
    columnsRoundPoints[4] = 50;
    columnsRoundPoints[5] = 60;
  }  
  // B, C, D, E, F, G
  else if ( numOfPlayers <= 64 ) {
    
    columnsRoundPoints[1] = 20; 
    columnsRoundPoints[2] = 30;
    columnsRoundPoints[3] = 40;
    columnsRoundPoints[4] = 50;
    columnsRoundPoints[5] = 60;
    columnsRoundPoints[6] = 70;
    
  }
  else if ( numOfPlayers <= 128 ) {
    
    columnsRoundPoints[1] = 20; 
    columnsRoundPoints[2] = 30;
    columnsRoundPoints[3] = 40;
    columnsRoundPoints[4] = 50;
    columnsRoundPoints[5] = 60;
    columnsRoundPoints[6] = 70;
    columnsRoundPoints[6] = 80;
  }
}

function tournGradeC(numOfPlayers)
{
  // A column is 1st round, so no points
  columnsRoundPoints[0+startColumn] = 0 ;
  // B , C , D
  if ( numOfPlayers <= 8 ) {
    columnsRoundPoints[1] = 10; 
    columnsRoundPoints[2] = 20;
    columnsRoundPoints[3] = 30;
  } 
  // B, C, D , E
  else if ( numOfPlayers <= 16 ) {
    columnsRoundPoints[1] = 10; 
    columnsRoundPoints[2] = 20;
    columnsRoundPoints[3] = 30;
    columnsRoundPoints[4] = 40;
  }
  // B, C, D, E, F
  else if ( numOfPlayers <=32 ) {
    columnsRoundPoints[1] = 10; 
    columnsRoundPoints[2] = 20;
    columnsRoundPoints[3] = 30;
    columnsRoundPoints[4] = 40;
    columnsRoundPoints[5] = 50;
  }  
  // B, C, D, E, F, G
  else if ( numOfPlayers <= 64 ) {
    
    columnsRoundPoints[1] = 10; 
    columnsRoundPoints[2] = 20;
    columnsRoundPoints[3] = 30;
    columnsRoundPoints[4] = 40;
    columnsRoundPoints[5] = 50;
    columnsRoundPoints[6] = 60;
    
  }
  else if ( numOfPlayers <= 128 ) {
    
    columnsRoundPoints[1] = 10; 
    columnsRoundPoints[2] = 20;
    columnsRoundPoints[3] = 30;
    columnsRoundPoints[4] = 40;
    columnsRoundPoints[5] = 50;
    columnsRoundPoints[6] = 60;
    columnsRoundPoints[7] = 70;
  }
}

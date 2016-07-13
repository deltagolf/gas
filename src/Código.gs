function main(){
  dataCrossFill();
  test();
}


function dataCrossFill() {
  
  var bookID = "14uYlNFllcNc1L2m48YsG-5aYXizdbmTPTaIZsX8iY4Q", mediaSheet = "MEDIA_FILES", mergeSheet = "DATA_MERGE", peopleSheet = "PEOPLE";
  
  var merge = new Merge(bookID, mergeSheet);
  var media = new Media(bookID, mediaSheet);
  var people = new People(bookID, peopleSheet);
  
  var mediaData = media.contents;
  var peopleData = people.contents;
  var mergeData = merge.contents;

  people.appearance(mergeData).addAppearance();
  merge.addedMedia(mediaData).updateMedia();
  merge.addedPeople(peopleData).updatePeople();
 
}

/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/


function test() {
  
  //application/vnd.openxmlformats-officedocument.wordprocessingml.document
  //, type = "application/vnd.oasis.opendocument.text", array;
  //array = downloadAsType(folderID, type);
  var sourceFolderID = "0B8avh_qu7PYJTEZjRnMtcS1pY1U", masterFolderID = "0B8avh_qu7PYJMjN4aGdqRndXYXM", templateID = "1p9pGtcXJXGArV41hwBCtNIxToG8AJAgOAQa7QaynqbU"; 
  var id = docFormat(masterFolderID, sourceFolderID, templateID);
  id = downloadLinks(id);
  Logger.log(id);
}

function docFormat(masterFolderID, sourceFolderID, templateID){
  
  //Declare constants (test files, folders, etc)
  //Spreadsheet info
  var bookID = "14uYlNFllcNc1L2m48YsG-5aYXizdbmTPTaIZsX8iY4Q", mediaSheet = "MEDIA_FILES", mergeSheet = "DATA_MERGE", peopleSheet = "PEOPLE";
  //General variables
  var files, id, file, name, body, paragraphsNumber, paragraph, text, i, k = 1, j, n, m, r, s,
      bodiesLength, attributes, mergeIDsLength, flag, table, tableElement, row, numRows, header, interviewArray, interviewObject, mergeIDS, numChildren, numChildrenIndex, masterTables, numTables, newRow, people, tablePeople, media, tableMedia,
      masterDoc, docCell, masterDocId, masterDocBody, masterDocHeader, masterDocFooter, placeHolder, subDocArray;
  var placeHolders = ["AMMEND", "INTERVIEW_ID", "TITLE", "DATE", "LOCATION", "VERSION", "KEYWORDS"];
  
  //Array with the bodies objects from the source documents
  var bodiesArray = [];
  //Get folder object for master documents
  var folder = DriveApp.getFolderById(masterFolderID);

  //Get DATA_MERGE info
  //Array of all interviews data;
  var sheetValues = SpreadsheetApp.openById(bookID).getSheetByName(mergeSheet).getDataRange().getValues();
  //Get number of elements
  var dataMergeLength = sheetValues.length;
  //Loops through each interview
//  for(k; k < dataMergeLength; k++){
  for(k; k < 3; k++){
    //Sub-array with interview data
    interviewArray = sheetValues[k];
    interviewObject = toObject(interviewArray, sheetValues[0]);
    name = interviewObject.interview_id;
    if(interviewObject.files){
      //Create template clone and assign name
      masterDoc = DriveApp.getFileById(templateID).makeCopy(name, folder);
      masterDocId = masterDoc.getId();
      masterDocBody = DocumentApp.openById(masterDocId).getBody();
      masterDocHeader = DocumentApp.openById(masterDocId).getHeader();
      masterDocFooter = DocumentApp.openById(masterDocId).getFooter();
      m = 0;
      //Replace placeholders
      for(m; m < placeHolders.length; m++){
        placeHolder = placeHolders[m].toLowerCase();
        masterDocBody.replaceText("%" + placeHolders[m] + "%", interviewObject[placeHolder]);
      }
      masterDocBody.replaceText("%DURATION%", toHHMMSS(interviewObject.duration));
      masterDocFooter.replaceText("%SECTION%", "Testing Footer");
      masterTables = masterDocBody.getTables();
      numTables = masterTables.length;
      numChildrenIndex = 0;
      //Set people values
      people = JSON.parse(interviewObject.people);
      tablePeople = masterTables[3];
      r = 0;
      for(r; r < people.length; r++){
        newRow = tablePeople.appendTableRow();
        newRow.appendTableCell().setText(people[r].role);
        newRow.appendTableCell().setText(people[r].name +  " " + people[r].last_name);
        newRow.appendTableCell().setText(people[r].initials);
      }
      //Set media values
      media = JSON.parse(interviewObject.files);
      r = 0;
      var docsArray = [];
      var mergeIDs = [];
      for(r; r < media.length; r++){
        if(media[r].type == "V"){
          //For Video
          tableMedia = masterTables[4].getRow(0).getCell(0).getChild(1).asTable();
          newRow = tableMedia.appendTableRow();
          newRow.appendTableCell().setText(media[r].fileName);
          newRow.appendTableCell().setText(toHHMMSS(media[r].duration));
          newRow.appendTableCell().setText(media[r].from);
          newRow.appendTableCell().setText(media[r].to);
          
        } else if (media[r].type == "A"){
          //For Audio
          tableMedia = masterTables[4].getRow(0).getCell(1).getChild(1).asTable();
          newRow = tableMedia.appendTableRow();
          newRow.appendTableCell().setText(media[r].fileName);
          newRow.appendTableCell().setText(toHHMMSS(media[r].duration));
          newRow.appendTableCell().setText(media[r].from);
          newRow.appendTableCell().setText(media[r].to);
        } else if (media[r].type == "G"||media[r].type == "I"||media[r].type == "F"||media[r].type == "D"||media[r].type == "P"){
          //For documents and photos
          docsArray.push([media[r].type, media[r].fileName]);
          if(media[r].type == "G")
            mergeIDs.push([media[r].doc_id,media[r].linked_to]);
        }
      }
      docsArray.push(["G", name]);
      var googleDoc = docsArray.filter(function(i){
        if(i[0] === "G") 
          return true;
        else 
            return false;});
      var wordDoc = docsArray.filter(function(i){
        if(i[0] === "D") 
          return true; 
        else 
            return false;});

      var imgDoc = docsArray.filter(function(i){
        if(i[0] === "F") 
          return true; 
        else 
            return false;});
      var indesignDoc = docsArray.filter(function(i){
        if(i[0] === "I") 
          return true; 
        else 
            return false;});
      var pdfDoc = docsArray.filter(function(i){
        if(i[0] === "P") 
          return true; 
        else 
            return false;});
      docsArray = [googleDoc, wordDoc, indesignDoc, pdfDoc, imgDoc];
      var docsHeaders = ["Google Docs: ", "Word: ", "InDesign: ", "PDF: ", "Fotos: "];
      var docsParagraph;
      docCell = masterTables[7].getRow(0).getCell(1);
      for (var y = 0; y < docsArray.length; y++){
        if(docsArray[y].length > 0){
          var docText = docsHeaders[y];
          for each (var w = 0; w < docsArray[y].length; w++){
            docText += docsArray[y][w][1];
            docText +="; ";
          }
          docCell.appendParagraph(docText);
        }
      }
      
      //Append the documents' texts
      var numFiles = mergeIDs.length;
      i = 0;
      //Get all the documents bodies
      for(i; i < numFiles; i++){
        var arrayRow = [];
        arrayRow.push(mergeIDs[i][0]);
        var document = DocumentApp.openById(mergeIDs[i][0]);
        arrayRow.push(document.getName());
        arrayRow.push(document.getBody());
        bodiesArray.push(arrayRow);
      }
      //Iterate through bodies and append tables and contents
      n = 0;
      //New method to find linked files
      media.findLinkedFile = function(name){
        var item = this, answer;
        var length = item.length;
        for(var i = 0; i < length; i++){
          if(name == item[i].fileName)
            answer = item[i].linked_to;
        }
        if(answer)
          return answer;
        else
          return null;
      };
      var counter = 1;
      for(n; n < numFiles; n++){
        body = bodiesArray[n][2];
        var fileName = bodiesArray[n][1];
        var audioDoc = media.findLinkedFile(fileName);
        var videoFile = media.findLinkedFile(audioDoc);
        var textTableText = audioDoc + " / " +videoFile + "\n" + interviewObject.title;
        var textTableCell = [[textTableText]];
        masterDocBody.appendPageBreak();
        masterDocBody.appendTable().appendTableRow().appendTableCell(textTableText);
        var textTable = masterDocBody.appendTable();
        var bodyLength = body.getNumChildren();
        for(j = 0; j < bodyLength; j++){
          var content = body.getChild(j).getText();
          if(content){
            var textTableRow = textTable.appendTableRow();
            textTableRow.appendTableCell().setWidth(50);
            textTableRow.appendTableCell().setWidth(28);
            textTableRow.appendTableCell().setText(content);
            //Set content
            //textTableRow.appendTableCell().setWidth(28).setText(counter++);
            //No content
            textTableRow.appendTableCell().setWidth(28);
          }
        }
      }
    }
  }
  return masterDocId;
}

/*                                                                        OBJECTS                                                                           */
/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/


var Merge = function(id, sheet){
    this.id = id;
    this.sheet = sheet;
    this.contents;
    this.data = function(){
      var data = SpreadsheetApp.openById(this.id).getSheetByName(this.sheet).getDataRange().getValues();
      this.contents = data;
      return this;
    };
    this.addedMedia = function(media){
      var lengthMedia, lengthData, i = 1, k = 1,  mediaData, mediaItem, dataItem, data, mediaRow;
      data = this.contents;
      lengthMedia = media.length;
      lengthData = data.length;
      if(lengthData > 1 && lengthMedia > 1){
        for(k; k< lengthData; k++){
          i = 1;
          dataItem = data[k];
          for(i; i<lengthMedia; i++){
            mediaRow = media[i];
            mediaItem = {};
            if(dataItem[0] === mediaRow[2]){
              mediaItem.fileName = mediaRow[0],
                mediaItem.type = mediaRow[1],
                  mediaItem.interview = mediaRow[2],
                    mediaItem.duration = mediaRow[3],
                      mediaItem.from = mediaRow[4],
                        mediaItem.to = mediaRow[5],
                          mediaItem.linked_to = mediaRow[6],
                            mediaItem.doc_id = mediaRow[7];
              
              if(dataItem[5]){
                mediaData = dataItem[5];
              } else {
                mediaData = [];
              }
              if(k==2)
              mediaData.push(mediaItem);
              dataItem[5] = mediaData;
              data[k] = dataItem;
            }
          }
          mediaData = data[k][5];
          if(mediaData != ""){
            data[k][5] = JSON.stringify(mediaData);
          } else {
            data[k][5] = null;
          }
        }
      } 
      this.contents = data;
      return this;
    };
    this.updateMedia = function(){
      var addedMedia = this.contents;
      SpreadsheetApp.openById(this.id).getSheetByName(this.sheet).getDataRange().setValues(addedMedia);
      return addedMedia;
    };
    this.addedPeople = function(people){
      var lengthPeople, lengthData, k = 1,  peopleData, dataItem, data, peopleRow, interviewers, participants, arr1, arr2;
      data = this.contents;
      lengthPeople = people.length;
      lengthData = data.length;
      if(lengthData > 1 && lengthPeople > 1){
        for(k; k< lengthData; k++){
          dataItem = data[k];
          arr1 = dataItem[10];
          arr1 = arr1.split(";");
          arr2 = dataItem[11];
          arr2 = arr2.split(";");
          interviewers = getPeopleData(arr1, "I", people);
          participants = getPeopleData(arr2, "P", people);
          peopleData = interviewers.concat(participants);
          dataItem[13] = peopleData;
          data[k] = dataItem;
          peopleData = data[k][13];
          if(peopleData != ""){
            data[k][13] = JSON.stringify(peopleData);
          } else {
            data[k][13] = null;
          }
        }
        
      }
      this.contents =data;
      return this;
    };
    this.updatePeople = function(){
      var addedPeople = this.contents;
      SpreadsheetApp.openById(this.id).getSheetByName(this.sheet).getDataRange().setValues(addedPeople);
      return addedPeople;
    };
  };

var People = function(id, sheet){
  this.id = id;
  this.sheet = sheet;
  this.contents;
  this.data = function(){
    var data = SpreadsheetApp.openById(this.id).getSheetByName(this.sheet).getDataRange().getValues();
    this.contents = data;
    return this;
  };
  this.appearance = function(merge){
    var lengthMerge, lengthData, i = 1, k = 1,  mergeData, mergeItem, dataItem, data, person, role, mergeRow, interview;
    data = this.data();
    lengthMerge = merge.length;
    lengthData = data.length;
    if(lengthData > 1 && lengthMerge > 1){
      for(k; k< lengthData; k++){
        i = 1, mergeData = [];
        dataItem = data[k];
        for(i; i<lengthMerge; i++){
          mergeRow = merge[i];
          person = dataItem[4];
          role = dataItem[0];
          interview = findPeople(person, role, mergeRow);
          if(interview){
            mergeData.push(interview);
          }
        }
        mergeData = JSON.stringify(mergeData);
        dataItem[5] = mergeData;
        data[k] = dataItem;
      }
    } 
    this.contents = data;
    return this;
  };
  this.addAppearance = function(){
    var addedAppearance = this.contents;
    SpreadsheetApp.openById(this.id).getSheetByName(this.sheet).getDataRange().setValues(addedAppearance);
    return addedAppearance; 
  }
};

var Media = function(id, sheet){
  this.id = id;
  this.sheet = sheet;
  this.contents;
  this.data = function(){
    var data = SpreadsheetApp.openById(this.id).getSheetByName(this.sheet).getDataRange().getValues();
    this.contents = data;
    return this;
  };
};





function findPeople(person, role, array){
  var column, peopleData, interview, length, i = 0;
  role = role.trim().toUpperCase();
  if (role == "I"){
    column = 10;
  } else if (role == "P"){
    column = 11;
  } else {
    column = null;
  }
  if(column){
    //Do stuff
    interview = array[0];
    peopleData = array[column];
    peopleData = peopleData.split(";");
    length = peopleData.length;
    for(i; i < length; i++){
      if(peopleData[i] === person){
        return interview;
      }
    }
  } 

}

function getPeopleData(arr, role, peopleMatrix){
  var length, matrixLength, matrixRow, i = 0, k, person, personObject, objectsArray = [];
  length = arr.length;
  matrixLength = peopleMatrix.length
  if(length > 0){
    for(i; i<length; i++){
      k = 1, personObject = {};
      person = arr[i];
      for(k; k < matrixLength; k++){
        matrixRow = peopleMatrix[k];
        if(person == matrixRow[4]&&role == matrixRow[0]){
          personObject.role = matrixRow[0],
            personObject.name = matrixRow[1],
              personObject.last_name = matrixRow[2],
                personObject.position = matrixRow[3],
                  personObject.initials = matrixRow[4];
          objectsArray.push(personObject);
        }
      }
    }
    return objectsArray;
  }
}


Array.prototype.clean = function(deleteValue) {
  for (var i = 0; i < this.length; i++) {
    if (this[i] == deleteValue) {         
      this.splice(i, 1);
      i--;
    }
  }
  return this;
};

/*                                                                UTILITY FUNCTIONS                                                                         */
/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/
/************************************************************************************************************************************************************/

function toHHMMSS(secs){
  var time = new Date(secs * 1000);
  var hh = time.getHours() - 1;
  hh = zeroFill(hh, 2) + ":";
  var mm = time.getMinutes();
  mm = zeroFill(mm, 2) + ":";
  var ss = zeroFill(time.getSeconds(), 2);
  time = hh+mm+ss;
  return time;
}

function zeroFill( number, width )
{
  width -= number.toString().length;
  if ( width > 0 )
  {
    return new Array( width + (/\./.test( number ) ? 2 : 1) ).join( '0' ) + number;
  }
  return number + ""; // always return a string
}


function downloadAllItemsAsType(folderID, type){
  var urlArray, id, url, name, item, files, file, row;
  files = Drive.Children.list(folderID).items;
  urlArray = [];
  for(var item in files){
    id = files[item].id;
    file = Drive.Files.get(id);
    name = file.title;
    url = file.exportLinks[type];
    row = [id, name, url];
    urlArray.push(row);
  }
  return urlArray;
}

function downloadLinks(fileID){
  var links = Drive.Files.get(fileID).exportLinks;
  return links;
}

function toObject(arr, keys) {
  var rv = {};
  for (var i = 0; i < arr.length; ++i)
    rv[keys[i].toLowerCase()] = arr[i];
  return rv;
}

/** 
 * Author:   Benjamin Borden
 * Created:  01/20/2023
 * Modified: 01/20/2023
 */

var getDoc = DocumentApp.openByUrl("")
var getDocBody = getDoc.getBody();
var getDocHeader = getDoc.getHeader();

var cachedImages;
var document;
var spreadsheet;
var sheet;
var cellRange; 
var nameOfCurrentSheet;

function publishToPlayers(){

  spreadsheet = SpreadsheetApp.getActive();
  sheet = spreadsheet.getActiveSheet();
  cellRange = sheet.getActiveRange();

  cachedImages = new Map();

  var firstColumn = cellRange.getColumn();
  var firstRow = cellRange.getRow();

  const numRows = cellRange.getNumRows();
  const numColumns = cellRange.getNumColumns();

  const sheetType = sheet.getRange(1,1).getValue();

  nameOfCurrentSheet = sheet.getSheetName();
  let sheetElem = findElementWithText(nameOfCurrentSheet);
  if(sheetElem === null){
    var newSheetElem = getDocBody.appendParagraph(nameOfCurrentSheet);
    newSheetElem.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    sheetElem = newSheetElem;
  }
  const cellRangeValues = cellRange.getValues();
  const cellRangeBackgrounds = cellRange.getBackgrounds();

  for(var i = firstColumn; i<firstColumn+numColumns;i++){

    const identifier = sheet.getRange(2,i).getValue();
    for(var o = firstRow; o<firstRow+numRows;o++){
      Logger.log("Column: "+i+"\nRow: "+o);
      var currentCell = sheet.getRange(o,i);
      if(currentCell.getBackground() !== "#fffecf"){
        Logger.log("Improper Highlighting");
        Logger.log(currentCell.getBackground());
        Logger.log(currentCell.getValue());
        return
      }

      const isImage = currentCell.getValue()?.valueType === SpreadsheetApp.ValueType.IMAGE;

      switch(sheetType){
        case "Column":
          Logger.log("Column case begun");
          
          let column = findElementWithText(identifier);

          if(column === null){
            var newParagraph = getDocBody.insertParagraph(getDocBody.getChildIndex(sheetElem)+1,identifier);
            newParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
            column = newParagraph
          } 
          
          Logger.log("column:"+column);
          //Logger.log("child index:"+getDocBody.getChildIndex(column));
          if(isImage){ 
            const inlineImg = column.asParagraph().appendInlineImage(retrieveImageBlob(currentCell));
            const width = (300/inlineImg.getHeight()) * (inlineImg.getWidth());
            inlineImg.setHeight(300).setWidth(width);
          }else{
            if(currentCell.getValue() === ""){
              break;
            }
            var newListItem = addValueAtIndex(getDocBody.getChildIndex(column)+1,currentCell);
            if(newListItem.getText() !== ""){
              newListItem.setText(sheet.getRange(o,1).getValue()+": "+newListItem.getText());
            }
            const nextSibling = newListItem?.getNextSibling();
            if(nextSibling !== null && newListItem.getNextSibling().getType() === DocumentApp.ElementType.LIST_ITEM){
                newListItem.setListId(nextSibling);
            }
          }
          break;
        case "Row":
          break;
      }

      currentCell.setBackground("#e1ffe3");

    }
    updateHeader("New information about "+identifier+"!");
  }

  
}

function updateHeader(text){
  var date = Utilities.formatDate(new Date(), "GMT-5", "MM/dd HH:mm")
  getDocHeader.insertParagraph(0,"["+date+"] "+text).setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  getDocHeader.removeChild(getDocHeader.getChild(3));
}
function addValueAtIndex(index, cell){
  const value = cell.getValue();
  var rtrn;
  if(value?.valueType === SpreadsheetApp.ValueType.IMAGE){

  }else{
    rtrn = getDocBody.insertListItem(index,value).setGlyphType(DocumentApp.GlyphType.BULLET);
  }

  return rtrn;
}
function findElementWithText(text){
  let size = getDocBody.getNumChildren();
  for(var counter = 0; counter < size; counter++){
    var elem = getDocBody.getChild(counter);
    if(elem.getType() === DocumentApp.ElementType.PARAGRAPH){
      elem = elem.asParagraph();
      if(elem.getText() === text){
        return elem;
      }
    } else if(elem.getType() === DocumentApp.ElementType.LIST_ITEM){
      elem = elem.asListItem();
      if(elem.getText() === text){
        return elem;
      }
    }
  }
  return null;
}
function retrieveImageBlob(cell){
  var res;
  if(cachedImages.has(nameOfCurrentSheet)){
    res = cachedImages[nameOfCurrentSheet];
  }else{
    res = DocsServiceApp.openBySpreadsheetId(spreadsheet.getId()).getSheetByName(nameOfCurrentSheet).getImages();
  }
  Logger.log(res);
  for(img of res){
    Logger.log(img)
    if(img.range.col === cell.getColumn() && img.range.row === cell.getRow()){
      return img.image.blob;
    }
  }
  return null;
}

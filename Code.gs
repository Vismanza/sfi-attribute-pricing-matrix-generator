function onOpen() {
     SpreadsheetApp.getUi().createMenu('Salesforce Industries')
         .addItem('Generate ABP Matrix','generateABP')
         .addSeparator()
         .addItem('Append to ABP Matrix','appendABP')
         .addToUi()
}
function appendABP() {
generateABP('Append');
}

function generateABP(mode) {

var start = new Date();

   var sheet = SpreadsheetApp.getActive();


// Get Product Array
 var loadProductRange = sheet.getSheetByName('Products').getRange(2,1,sheet.getSheetByName('Products').getLastRow()-1,sheet.getSheetByName('Products').getLastColumn()).getValues().filter(function(item){return item[0] === true;});

 //console.log(loadProductRange);


// Get Attribute Array
 var loadAttRange = sheet.getSheetByName('Attributes').getRange(1,1,1,sheet.getSheetByName('Attributes').getLastColumn()).getValues();
 //console.log(loadAttRange[0]);

 var attRange = loadAttRange.join().split(',').filter(Boolean);

// Calculate number of attributes
 var attCount =  attRange.length;
 console.log("Attribute count: " + attCount );
 

 var loadAttData = sheet.getSheetByName('Attributes').getRange(1,1,sheet.getSheetByName('Attributes').getLastRow(),attCount).getValues();

//console.log(loadAttData);

// Transpose attributes data into rows
 var attData = transpose(loadAttData);
 //console.log(attData);

// Trying for loop instead of map to loop through attribute value arrays

// for first attribute
var attFirst = attData[0].filter(String);
var attFirstLength = attFirst.length;
console.log(attFirstLength - 1 + " attribute values for " + attFirst[0] + " (first attribute)");

attSetData = [];
comboArray = [];
newComboArray = [];
// create initial work set
for (l = 1; l < attFirstLength; l++)
  {
    attSetData.push([attFirst[0],attFirst[l]]);
  }
 //console.log("Initial work set (attSetData): ");
 //console.log(attSetData);

// loop per additional attibute
for (k = 1; k < attCount; k++) 
  {
    // Define number of attribute values for current attribute
    var trimmedAttData = attData[k].filter(String);
    var attLength = trimmedAttData.length;
    console.log(attLength - 1 + " attribute values for " + trimmedAttData[0]);
    
    //loop per attribute value
      for (l = 1; l < attLength; l++)
      {
          valueSetData = attSetData.map(
            function(row)
            {
              return [row[0] + ";" + trimmedAttData[0],row[1] + ";" + trimmedAttData[l]];
            }
          );
          //console.log("Per attribute value (valueSetData): ");
          //console.log(valueSetData);
          //push each array in valueSetData
            lengthValueSetData = valueSetData.length;
            for (p = 0; p < lengthValueSetData; p++)
              { 
                comboArray.push(valueSetData[p]);
              }
          
          //console.log("Per attribute value (comboArray): ");
          //console.log(comboArray);
          
      }
      //attSetData = comboArray;
      var attsSoFar = loadAttRange[0].slice(0,k+1);
      console.log(attsSoFar);
      filterIndAttValues = loadAttRange[0].slice(0,k+1).join(";");
      var finalIndAttArray = comboArray.filter(function(item){return item[0] === filterIndAttValues;});
      //console.log(finalIndAttArray);
      attSetData = finalIndAttArray;
  }

//filterAttValues = loadAttRange[0].flat()
filterAttValues = loadAttRange[0].join(";");

var finalAttArray = comboArray.filter(function(item){return item[0] === filterAttValues;});


console.log("Combo'd Attributes: " + filterAttValues);
//console.log("Final attribute combo array:");
//console.log(finalAttArray);

// Add product config to array

lengthLoadProductRange = loadProductRange.length;

workingPricedArray = [];
finalPricedArray = [];

// Loop per product config
for (q = 0; q < lengthLoadProductRange; q++)
      {
       
        workingPricedArray = finalAttArray.map(
          function(attCombo)
        {
          return[
              loadProductRange[q][1],
              loadProductRange[q][2],
              attCombo[0],
              attCombo[1],
              loadProductRange[q][3],
              loadProductRange[q][4],
              loadProductRange[q][5],
              loadProductRange[q][6],
              loadProductRange[q][7],
              loadProductRange[q][8],
              loadProductRange[q][9],
              loadProductRange[q][10],
              loadProductRange[q][11],
              loadProductRange[q][12]
              ];
        }
        );
        //console.log(workingPricedArray);

        lengthWorkingPricedArray = workingPricedArray.length;
        // Loop per combo record for each product
          for (r = 0; r < lengthWorkingPricedArray; r++)
          { 
            finalPricedArray.push(workingPricedArray[r]);
          }
      }
      
//console.log(finalPricedArray);
calculatedPrices = finalPricedArray.map(calcPrices);

console.log("Priced rows: " + calculatedPrices.length);

if(!mode){

  mode = "Generate"

  // Clear sheet
  var dataSheet = sheet.getSheetByName('ABP Data').clear();
 
  // Create Headings
  sheet.getSheetByName('ABP Data').getRange(1,1).setValue('Source Product Name');
  sheet.getSheetByName('ABP Data').getRange(1,2).setValue('Source Product Code');
  sheet.getSheetByName('ABP Data').getRange(1,3).setValue('Characteristic Name');
  sheet.getSheetByName('ABP Data').getRange(1,4).setValue('Characteristic Value');
  
  sheet.getSheetByName('ABP Data').getRange(1,5).setValue('Recurring Price');
  sheet.getSheetByName('ABP Data').getRange(1,6).setValue('Recurring Cost');
  
  sheet.getSheetByName('ABP Data').getRange(1,7).setValue('One Time Price');
  sheet.getSheetByName('ABP Data').getRange(1,8).setValue('One Time Cost');
  
  sheet.getSheetByName('ABP Data').getRange(1,9).setValue('Usage Price');
  sheet.getSheetByName('ABP Data').getRange(1,10).setValue('Usage Cost');
  sheet.getSheetByName('ABP Data').getRange(1,11).setValue('Charge Measurement'); 

 // Place array in sheet
 sheet.getSheetByName('ABP Data').getRange(2,1,calculatedPrices.length,calculatedPrices[0].length).setValues(calculatedPrices);

}
else if (mode == 'Append'){

 // Append array in sheet
 sheet.getSheetByName('ABP Data').getRange(sheet.getLastRow() + 1,1,calculatedPrices.length,calculatedPrices[0].length).setValues(calculatedPrices);

}


// log and stats
var end = new Date();
var executiontime = ((end - start) / 1000).toFixed(2) + "s"
sheet.getSheetByName('Stats').getRange(sheet.getLastRow() + 1,1).setValue(Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy HH:mm:ss"));
sheet.getSheetByName('Stats').getRange(sheet.getLastRow() + 0,2).setValue(attCount);
sheet.getSheetByName('Stats').getRange(sheet.getLastRow() + 0,3).setValue(calculatedPrices.length);
sheet.getSheetByName('Stats').getRange(sheet.getLastRow() + 0,4).setValue(executiontime);
sheet.getSheetByName('Stats').getRange(sheet.getLastRow() + 0,5).setValue(mode);
sheet.getSheetByName('Stats').getRange(sheet.getLastRow() + 0,6).setValue(filterAttValues);


SpreadsheetApp.getActive().toast("You priced " + calculatedPrices.length + " combinations from " + attCount + " attributes in " + executiontime);

}  

function randomInteger(min, max) {
  if (!min || !max)
  {
    return 0
  }
  else {
  return Math.floor(Math.random() * (max - min + 1)) + min;
  }
}

function calcCost(price, margin) {
  if (price == 0){return 0;}
  else if (margin == 'None'){return 0;}
  else if (margin == 'Low'){return price - (price * randomInteger(80,100)/100)}
  else if (margin == 'Med'){return price - (price * randomInteger(50,80)/100)}
  else if (margin == 'High'){return price - (price * randomInteger(10,50)/100)}
  else {return 0}
}

function transpose(a)
{
  return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
}

function calcPrices(row)
{
 
  var recurringCharge = randomInteger(row[8],row[9]).toFixed(2);
  var recurringCost = calcCost(recurringCharge,row[10]).toFixed(2);
  var oneTimeCharge = randomInteger(row[11],row[12]).toFixed(2);
  var oneTimeCost = calcCost(oneTimeCharge,row[13]).toFixed(2);
  var usageCharge = (randomInteger(row[5]*100000,row[6]*100000)/100000).toFixed(5);
  var usageCost = (calcCost(usageCharge,row[7])).toFixed(5);
  if (row[4] == ''){var chargeMeasurement = 'N/A'}
  else {var chargeMeasurement = row[4]}

  return [
    row[0],
    row[1],
    row[2],
    row[3],
    recurringCharge,
    recurringCost,
    oneTimeCharge,
    oneTimeCost,
    usageCharge,
    usageCost,
    chargeMeasurement
    ];
}






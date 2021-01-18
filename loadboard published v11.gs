var days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
var months = ['Jan','Feb','Ma','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
var usersPermitted = ['dispatch@martianexpress.us',]
//automate this
  var columnsLetters = {

    "LOAD#":'A',
    "TRUCK":'B',   
    "DRIVER":'D', 
    "PHONE":'E',
    "COVERED":'F',
    "BROKER":'K',
    "PU":'G',
    "DEL":'H',
    "MILES":'I',
    "RATE":'J',
    "DISPATCH":'L'

  }

//on install  
function onInstall(e) {
  onOpen(e);
}


function onOpen(e) {
  // var scriptID = ScriptApp.getScriptId(); alert(scriptID)
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [ 
    {name: 'Gross', functionName: 'Gross'},{name: 'New Day', functionName: 'newDay'}, { name: 'Locations', functionName: 'getLocations'},{ name: 'Night Board', functionName: 'nightShiftBoard'},{name: 'New Week', functionName: 'newWeek'} //,{name: 'Test 2', functionName: 'testMerge'}
  ];
  spreadsheet.addMenu('ZMXC', menuItems);
}

function getLocations(notes=false){
  try{
        var spreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loadboard");
        var lastRow = spreadSheet.getLastRow();
        
        

        var response = UrlFetchApp.fetch("https://api.samsara.com/fleet/vehicles/locations", {
        "headers": {
          "Authorization": "Bearer samsara_api_irWxh0GBjjVJfejxaDUmNLYAty0QaX"
        }
        });
        response = JSON.parse(response);
        var lngth = response.data.length;
        
        var response2 =UrlFetchApp.fetch("https://api.samsara.com/fleet/hos/clocks", {
          "headers": {
            "Authorization": "Bearer samsara_api_irWxh0GBjjVJfejxaDUmNLYAty0QaX"
          }
        });
        response2 = JSON.parse(response2);
        var lngth2 = response2.data.length;

        var bgrry = {},bgrry2 = {}
        for(var i=0;i<lngth;i++){
            // alert(i)
            var dt = response.data[i];
            
            bgrry[dt.name] = [dt.location.reverseGeo.formattedLocation, dt.location.speed];
        }
        
        for(var i=0;i<lngth2;i++){
          
            var dt = response2.data[i];
            var dtcycle = dt.clocks.cycle
          
            bgrry2[dt.driver.name.toUpperCase()]=(dtcycle.cycleRemainingDurationMs/3600000).toFixed()+" ("+(dtcycle.cycleTomorrowDurationMs/3600000).toFixed()+") " //= [dt.location.reverseGeo.formattedLocation, dt.location.speed];
        }
        // alert(JSON.stringify(bgrry2))
        // alert(JSON.stringify(bgrry))
        for(var i = lastRow-1; i>1 && !isDate(spreadSheet.getRange("A"+i).getValue());i--){


          try{
            var rnge = spreadSheet.getRange(columnsLetters["TRUCK"]+i)
            var rngeVal = rnge.getValue();

            if(rnge.isPartOfMerge()){
              continue;
            }

            rngeVal = isNaN(rngeVal)? false: rngeVal
            var driverArray = spreadSheet.getRange(columnsLetters["DRIVER"]+i).getValue().replace('no expedited loads','').toUpperCase().replace('SOLO','').replace('TEAM','').replace(/ *\([^)]*\) */g, "").split(" / ");
            
            rngeValBool = rngeVal
            if(SpreadsheetApp.getActive().getName() == 'Loadboard MRPP'){
              rngeVal+=" Martian Express CO"        
            }
            
            if(rngeValBool && rngeVal in bgrry &&  !spreadSheet.getRange(columnsLetters["COVERED"]+i).getValue().toUpperCase().includes('COVERED')){
              var loc = bgrry[rngeVal][0]
              var loc = loc.split(",");
              var loclen = loc.length;
              


              var hrsofsrv = bgrry2[driverArray[0]]
              hrsofsrv+= (1 in driverArray)? " / "+bgrry2[driverArray[1]]:"";
              if(!hrsofsrv.includes("undefined")){
                spreadSheet.getRange(columnsLetters["MILES"]+i).setValue(hrsofsrv) //loc[loclen-2] +", "+loc[loclen-1]
                spreadSheet.getRange(columnsLetters["PU"]+i).setValue(loc[loclen-2] +", "+loc[loclen-1]) //loc[loclen-2] +", "+loc[loclen-1] bgrry[rngeVal][0]
                spreadSheet.getRange("A"+i).setValue(parseInt(bgrry[rngeVal][1])+" mi/h ") //loc[loclen-2] +", "+loc[loclen-1]
              }
            }
            if(notes){
              spreadSheet.getRange(columnsLetters["RATE"]+i+":"+columnsLetters["BROKER"]+i).setValues(notes[spreadSheet.getRange(columnsLetters["TRUCK"]+i).getValue()])
            }
          }catch(e){ 
          }
        }

        var today = new Date();
        var h = today.getHours();
        var m = today.getMinutes();
        var s = today.getSeconds();
  }catch(e){
    alert("Error getting Location ")
  }
 
}


function testMerge(){
  var spreadsheet = SpreadsheetApp.getActive();

  
   spreadsheet.getRange('A80').setValue(new Date());
 

}

function valid(){
  if(Session.getActiveUser().getEmail() == 'dispatch@martianexpress.us'){
    return true;
  }else{
    alert("You can not use this function")
    return false;
  }
}
function Gross(){
  // if(valid()!=true) return
  
  getGross(); // calls setGross
}


function setGross(bigArray,currentDay){
  try{
  
 
          var daysRow=1;
          var startDayColumn = 3;
          var endDayColumn = 'I';
          var DriverColumn = 'A';
          var FirstDriverRow = 3
          var stepPropertiesOfLoad = 5
          var spreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gross");
          var lastRow = spreadSheet.getLastRow();
          var lastColumn = spreadSheet.getLastColumn();
        

          for(j=startDayColumn;j<lastColumn;j++){
    
            var colVal = spreadSheet.getRange(daysRow,j).getValue()
          
            if (!(isDate(colVal))) continue;
            var colValTimeDay = colVal.getTime()
          
            if(!(colValTimeDay in bigArray)) continue;
          
            for(var i =FirstDriverRow; i<=lastRow;i+=1){ //stepPropertiesOfLoad
              var tmpDriver=spreadSheet.getRange(DriverColumn+i).getValue()
            
          
              
              if(tmpDriver=="" ){
                continue;
              }
              
              try{
              tmpDriver =  tmpDriver.match(/\d+/)[0]
              if(tmpDriver.length == 0){
                continue;
              }
              }
                catch(er){
        
                }
              
              
              if(tmpDriver in bigArray[colValTimeDay]){
                // first we set considering one covered per day
                
                for(k=0;k<bigArray[colValTimeDay][tmpDriver].length;k++){
                  if(k==0){
                    
                    spreadSheet.getRange(i,j).setValue(bigArray[colValTimeDay][tmpDriver][k]["LOAD#"]);        
                    spreadSheet.getRange(i+1,j).setValue(bigArray[colValTimeDay][tmpDriver][k]["PU"]);
                    spreadSheet.getRange(i+2,j).setValue(bigArray[colValTimeDay][tmpDriver][k]["DEL"]);
                    spreadSheet.getRange(i+3,j).setValue(bigArray[colValTimeDay][tmpDriver][k]["RATE"]);
                    spreadSheet.getRange(i+4,j).setValue(bigArray[colValTimeDay][tmpDriver][k]["MILES"]);
                  }else{
                    var oldLoad = spreadSheet.getRange(i,j).getValue();
                      spreadSheet.getRange(i,j).setValue(oldLoad+"|"+bigArray[colValTimeDay][tmpDriver][k]["LOAD#"]); 
                    var oldpu = spreadSheet.getRange(i+1,j).getValue();
                      spreadSheet.getRange(i+1,j).setValue(oldpu+"|"+bigArray[colValTimeDay][tmpDriver][k]["PU"]);
                    var olddel = spreadSheet.getRange(i+2,j).getValue();
                      spreadSheet.getRange(i+2,j).setValue(olddel+"|"+bigArray[colValTimeDay][tmpDriver][k]["DEL"]);
                    var oldRate = spreadSheet.getRange(i+3,j).getValue() ;
                    var newRate =parseInt(oldRate)+parseInt(bigArray[colValTimeDay][tmpDriver][k]["RATE"]);
                      spreadSheet.getRange(i+3,j).setValue(newRate);
                    var oldMile = spreadSheet.getRange(i+4,j).getValue() ;
                    var newRate =parseInt(oldMile)+parseInt(bigArray[colValTimeDay][tmpDriver][k]["MILES"]);
                      spreadSheet.getRange(i+4,j).setValue(newRate);            
                  }
                }
              }
            }
            
          } 
  }catch(e){
    alert("Error setting  Gross ")
  }        
}


  
function getGross(){
  try{
        
          var spreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loadboard");
          var nameOfSheet = spreadSheet.getName();
          var startFrom = 'A1';
          var lastRow = spreadSheet.getLastRow();
          var bigArray=new Array(8);
          var currentDay = 0
          
          for(var i =1; i<=lastRow;i++){
            var AcellValueDay = spreadSheet.getRange('A'+i).getValue();
            var rngVal = spreadSheet.getRange(columnsLetters["COVERED"]+i).getValue().toUpperCase();
            if(isDate(AcellValueDay)){
              currentDay = AcellValueDay.getTime() //.getDay()
            
              bigArray[currentDay]=new Array();
        //      alert(1+JSON.stringify(coveredInstance))
            }else if ( rngVal.includes('COVERED') ){
              var driver = spreadSheet.getRange(columnsLetters["TRUCK"]+i).getValue();
              
              if(!(driver in bigArray[currentDay])){
                bigArray[currentDay][driver]=new Array();
              }
              var coveredInstance = {
                "LOAD#":spreadSheet.getRange(columnsLetters["LOAD#"]+i).getValue(),
                "PU":spreadSheet.getRange(columnsLetters["PU"]+i).getValue(),
                "DEL":spreadSheet.getRange(columnsLetters["DEL"]+i).getValue(),
                "MILES":spreadSheet.getRange(columnsLetters["MILES"]+i).getValue(),
                "RATE":spreadSheet.getRange(columnsLetters["RATE"]+i).getValue()
              }
              
              bigArray[currentDay][driver].push(coveredInstance)
        //      alert(2+JSON.stringify(coveredInstance))
            }else if (rngVal.includes('CANCELLED') || rngVal.includes('OFF') || rngVal.includes('OOS') || rngVal.includes('HOME') ){
              var driver = spreadSheet.getRange(columnsLetters["TRUCK"]+i).getValue();
              if(!(driver in bigArray[currentDay])){
                bigArray[currentDay][driver]=new Array();
              }
              var rt = spreadSheet.getRange(columnsLetters["RATE"]+i).getValue();
              var coveredInstance = {
                "LOAD#":rngVal+" "+spreadSheet.getRange(columnsLetters["LOAD#"]+i).getValue(),
                "PU":spreadSheet.getRange(columnsLetters["PU"]+i).getValue(),
                "DEL":spreadSheet.getRange(columnsLetters["DEL"]+i).getValue(),
                "MILES":0,
                "RATE":isNaN(rt)? "": rt
              }
        //      alert(3+JSON.stringify(coveredInstance))
              bigArray[currentDay][driver].push(coveredInstance)
        //      alert(JSON.stringify(bigArray))
            }
              
          }
        
        //  alert(JSON.stringify(coveredInstance))
          setGross(bigArray,currentDay)
  }catch(e){
    alert("Error getting  Gross ")
  } 
}


function newDay(){ 
  try{
          var spreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loadboard");
          var nameOfSheet = spreadSheet.getName();
          var startFrom = 'A1';
          var lastRow = spreadSheet.getLastRow();
          var lastColumn = spreadSheet.getLastColumn();
          

          
          var notes = false;
          if(lastRow>1){
            notes = myHideRow(spreadSheet,lastRow)
          }
        
          spreadSheet.getRange(lastRow+1,1,1,lastColumn).mergeAcross().setValue(dayDayMonth()).setFontWeight("bold");
          spreadSheet.getRange(lastRow+2,1,1,lastColumn).setValues(spreadSheet.getRange(1,1,1,lastColumn).getValues()).setFontWeight("bold").setBackground("#999999");;
          
          bigArrayDrivers = getDrivers();
          setDrivers(bigArrayDrivers, spreadSheet.getLastRow())
          getLocations(notes);
  }catch(e){
    alert("Error with creating  new Day ")
  } 
}

function dayDayMonth(daysago=0){
  if(daysago){
    var date = new Date();
    var now = new Date(date.getTime() - (daysago * 24 * 60 * 60 * 1000));
  } else{
    var now = new Date();
  }
  
  var day = days[ now.getDay() ];
  var dayM = now.getDate();
  var month = months[ now.getMonth() ];
  return day+" "+dayM+"-"+month
}

function isDate(date){ //isDate
  if (date instanceof Date && !isNaN(date.valueOf())){
    return true;
  }
  return false;
}


function getDrivers(){
  try{
          var spreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Driver List");
          var nameOfSheet = spreadSheet.getName();
          var lastRow = spreadSheet.getLastColumn();
          
          var driverColumnsNumber = {
            "DRIVER":'B', 
            "TRUCK":'B',
            "NUMBER":'G',
            "DISPATCH":'H'
          }
            
          for(var i=1;i<lastRow;i++){
            var AcellValueDay = spreadSheet.getRange(1,i).getValue();
        
            if(AcellValueDay == 'Truck'){
              driverColumnsNumber["DRIVER"]=i
            }else if(AcellValueDay == 'Driver Name'){
              driverColumnsNumber["TRUCK"]=i
            }else if(AcellValueDay == 'Number'){
              driverColumnsNumber["NUMBER"]=i
            }else if(AcellValueDay == 'Dispatch'){
              driverColumnsNumber["DISPATCH"]=i
            }
          }

          

          var startFrom = 'A1';
          var lastRow = spreadSheet.getLastRow();
          
          var currentDay = 0
          var dispatchers=[]
          for(var i =1;;i++){
            var AcellValueDay = spreadSheet.getRange('A'+i).getValue();
            
            if(AcellValueDay=="INACTIVE"){
              break;
            }
            else if(AcellValueDay=='#'){
              continue;

            }else{
              var dispatch = spreadSheet.getRange(i,driverColumnsNumber["DISPATCH"]).getValue();
              
              for(j=i;;j++){
                var nextdispatch = spreadSheet.getRange(j,driverColumnsNumber["DISPATCH"]).getValue();
                if(dispatch!=nextdispatch){
                  break;
                }
              }
        //      numberofdriverperdispatch=j-i
        //      alert(i+" and j is "+j+" and number of driv"+ numberofdriverperdispatch)
              
              if(!(dispatch in dispatchers)){
                dispatchers.push(dispatch)
        //        bigArrayDrivers[dispatch]={
        //          "DRIVER_NUMBER": spreadSheet.getRange(i,driverColumnsNumber["DRIVER"]+1,numberofdriverperdispatch,2).getValues(), 
        //          "TRUCK": spreadSheet.getRange(i,driverColumnsNumber["TRUCK"]+1,numberofdriverperdispatch,1).getValues()
        ////          "NUMBER": spreadSheet.getRange(i,driverColumnsNumber["NUMBER"],x,1).getValue(),
        //        };
        //        alert(JSON.stringify(bigArrayDrivers[dispatch]))
              }   
              i=j;
            }
            
          }
                bigArrayDrivers={
                  "DRIVER_NUMBER": spreadSheet.getRange(1
                                                        ,driverColumnsNumber["DRIVER"]+1,i-1,2).getValues(), 
                  "TRUCK": spreadSheet.getRange(1,driverColumnsNumber["TRUCK"]-1,i-1,1).getValues(),
                  "DISPATCH": spreadSheet.getRange(1,driverColumnsNumber["DISPATCH"],i-1,1).getValues()
        //          "NUMBER": spreadSheet.getRange(i,driverColumnsNumber["NUMBER"],x,1).getValue(),
                };
          
          return bigArrayDrivers;
  }catch(e){
    alert("Error getting drivers from driver list")
  }
}

function setDrivers(bigArrayDrivers, startfrom){
 try{
          var spreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loadboard");
          var lastRow = spreadSheet.getLastRow();
          var lastColumn = spreadSheet.getLastColumn(); 
          
          startfrom++;
          rowNumDriver = bigArrayDrivers["DRIVER_NUMBER"].length
          tmp = startfrom + rowNumDriver -1
         
          spreadSheet.getRange(columnsLetters["DRIVER"]+startfrom+':'+columnsLetters["PHONE"]+tmp ).setValues(bigArrayDrivers["DRIVER_NUMBER"]);
          spreadSheet.getRange(columnsLetters["TRUCK"]+startfrom+':'+columnsLetters["TRUCK"]+tmp ).setValues(bigArrayDrivers["TRUCK"]);
          spreadSheet.getRange(columnsLetters["DISPATCH"]+startfrom+':'+columnsLetters["DISPATCH"]+tmp ).setValues(bigArrayDrivers["DISPATCH"]);
          

          var lastRow = spreadSheet.getLastRow();
          for(i=startfrom-1,k=0;i<lastRow;i++,k++){
            
            if(spreadSheet.getRange(i,lastColumn).getValue() == 'Dispatch'){
              
              spreadSheet.getRange(i,1,1,lastColumn).mergeAcross().setValue(bigArrayDrivers["DISPATCH"][k+1]).setFontWeight("bold").setBackground("#b7b7b7"); //#999999 d9ead3 #b7b7b7
              
            }
          }
          var lastRow = spreadSheet.getLastRow();
          spreadSheet.getRange(lastRow,1,1,lastColumn).setValues(spreadSheet.getRange(1,1,1,lastColumn).getValues()).setFontWeight("bold").setBackground("#b7b7b7");  
  }catch(e){
    alert("Error inputing drivers into Loadboard")
  }
}
function myHideRow(spreadSheet,lastRow){ 
   try{ 
          var upperVal=0;
          var rangeToHide = 0;
          var bottomVal = lastRow-1
          var flagH = true;
          var notes=[]

        
          for(var i = lastRow-1; i>1 && !isDate(spreadSheet.getRange("A"+i).getValue());i--){
         
            var rnge = spreadSheet.getRange(columnsLetters["COVERED"]+i)
            var rngeVal = rnge.getValue().toUpperCase();
            // if(rnge.isPartOfMerge() ){ //|| rngeVal.includes('STATUS') 
            //   if(flagH){
            //     upperVal = i;
            //     spreadSheet.hideRows(upperVal+1, (bottomVal-upperVal));
            //   }
            //   // alert(i)
            //   bottomVal = i-1;
            //   flagH = true;
            //   continue;
            // }
            // after night shift used new day should be fine without merged cells
            if(flagH && (rngeVal.includes('COVERED') || rngeVal.includes('READY') || rngeVal.includes('STATUS') )){
            
              upperVal = i;
              spreadSheet.hideRows(upperVal+1, (bottomVal-upperVal));
              flagH=false;
            }

            if(!(rngeVal.includes('COVERED') || rngeVal.includes('CANCELLED'))){
            notes[spreadSheet.getRange(columnsLetters["TRUCK"]+i).getValue()] = spreadSheet.getRange(columnsLetters["RATE"]+i+":"+columnsLetters["BROKER"]+i).getValues()
            }
          }
          
          return notes
  } catch(e){
    alert("Error hiding rows")
  }
}

function nightShiftBoard(){ 
  try{
          var spreadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Loadboard");
          var lastRow = spreadSheet.getLastRow();
          var lastColumn = spreadSheet.getLastColumn();
          var upperVal=100;
        
          var bottomVal = lastRow-1
        
        

        
          for(var i = lastRow-1; i>1 && !isDate(spreadSheet.getRange("A"+i).getValue());i--){
            
            var rnge = spreadSheet.getRange(columnsLetters["COVERED"]+i)
            var rngeVal = rnge.getValue().toUpperCase();
            if(rnge.isPartOfMerge() ){ //|| rngeVal.includes('STATUS') 
              spreadSheet.deleteRow(i);
              continue;
            } else  if(rngeVal.includes('COVERED') || rngeVal.includes('CANCELLED') || rngeVal.includes('STATUS') ){
              spreadSheet.insertRowAfter(i);
              spreadSheet.getRange(i+1,1,1,lastColumn).setValues(spreadSheet.getRange(1,1,1,lastColumn).getValues()).setFontWeight("bold").setBackground("#b7b7b7"); 
              upperVal = i+2;
              // spreadSheet.sorRange(upperVal, (bottomVal-upperVal));
              break;
            }   
          }

          var conditions = ["COVERED","CANCELLED","HOLD","READY","DHING","UNLOADING","ETA","PU","CHECKED IN","OFF","OOS","HOME","ELD",""];
          var conditionsLetters = ["A","B","C","D","E","F","H","I","J","K","L","M","N","O"];
          var condlength = conditions.length;
          lastRow = spreadSheet.getLastRow();
          bottomVal = lastRow-1;
          var stsrange = spreadSheet.getRange(columnsLetters["COVERED"]+upperVal+":"+columnsLetters["COVERED"]+bottomVal);
          var statusValues = stsrange.getValues();
          var newstsVal = [];
          var stlen = statusValues.length;
        // alert(JSON.stringify(statusValues))
          for(var j =0;j<stlen;j++){
            var el = statusValues[j][0].toUpperCase();
            var fl = true
            // alert(el)
            for(var k=0;k<condlength;k++){
              if(el.includes(conditions[k])){
                newstsVal[j] = [conditionsLetters[k]+" "+statusValues[j][0]];
                fl=false;
                break;
              }
            }
            if(fl){
              newstsVal[j] = ["W"+statusValues[j][0]];
            }
              
          }
          stsrange.setValues(newstsVal)
          spreadSheet.getRange("A"+upperVal+":"+"L"+bottomVal).sort(6)

          statusValues = stsrange.getValues();
            for(var j =0;j<stlen;j++){
            
            
            
              statusValues[j][0] = statusValues[j][0].slice(2,);
            
              
          }
          stsrange.setValues(statusValues)
  } catch(e){
    alert("Error with night loadboard")
  }
}
 
// function rm(){alert(SpreadsheetApp.getActive().getName())}
function newWeek(){
  // DUPLICATE AND HIDE
  var newName = "Loadboard "+ dayDayMonth(7)+" - "+ dayDayMonth(1)
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadSheet = activeSpreadsheet.getSheetByName("Loadboard");

  var copy = spreadSheet.copyTo(activeSpreadsheet);
  copy.setName(newName) //.rename("Hello world");
  copy.hideSheet();


  // REMOVE OLD DATA AND ADD NEW DAY
  var lastRow = spreadSheet.getLastRow();
  spreadSheet.insertRowAfter(lastRow);
  spreadSheet.deleteRows(2,(lastRow-1));

  spreadSheet.getRange(2,1,1,spreadSheet.getLastColumn()).setBackground("white");
  spreadSheet.insertRowsAfter(2,1000);
  newDay();

}
 

function alert(text){
  
 var ui = SpreadsheetApp.getUi();
 ui.alert(text);
}

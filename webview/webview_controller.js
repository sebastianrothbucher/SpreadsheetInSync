/* 
   Licensed under the Apache License, Version 2.0 (the "License"); you may not
   use this file except in compliance with the License. You may obtain a copy of
   the License at
  
     http://www.apache.org/licenses/LICENSE-2.0
  
   Unless required by applicable law or agreed to in writing, software 
   distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
   WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
   License for the specific language governing permissions and limitations under
   the License. 
*/
angular.module("theplugin", []).controller("SheetController", function($scope, $http, $timeout, $window){
  $scope.userName=null;
  $scope.prevSince=0;
  $scope.sheets=[];
  $scope.sheetName=null;
  $scope.sheetData={};
  $scope.allSheetData={};
  $scope.sheetCols=[];
  $scope.sheetRows=[];
  $scope.filteredRows=[];
  $scope.updatedCells=[];
  $scope.selectedCell="A1";
  var colNames="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  function colToNum(addr){
    var col=addr.match(/[A-Z]+/)[0];
    var colnum=0;
    while(col.length>0){
      colnum=(colnum*26)+colNames.indexOf(col.substring(0, 1))+1;
      col=col.substring(1);
    }
    return colnum-1;
  };
  function numToCol(num){
    var col="";
    num++;
    while(num>0){
      if(num==26){
        col="Z"+col;
        break;
      }
      var ind=num%26;
      col=colNames.substring(ind-1, ind)+col;
      num-=ind;
      num/=26;
    }
    return col;
  };
  function changeAddr(addr, colDelta, rowDelta){
    // cols
    var col=addr.match(/[A-Z]+/)[0];
    var colIndex=$scope.sheetCols.indexOf(col);
    colIndex+=colDelta;
    col=$scope.sheetCols[Math.min($scope.sheetCols.length-1, Math.max(0, colIndex))];
    // row (consider filtering)
    var row=parseInt(addr.match(/[0-9]+/)[0]);
    var rowIndex=$scope.filteredRows.indexOf(row);
    rowIndex+=rowDelta;
    row=$scope.filteredRows[Math.min($scope.filteredRows.length-1, Math.max(0, rowIndex))];
    return ""+col+""+row;
  };
  function pollForChanges(){
    //console.log($scope.prevSince);
    $http({method: "GET", url: "../../../../_changes?since="+$scope.prevSince+"&feed=longpoll&include_docs=true"})
      .success(function(data, status){ 
        var rows=data.results;
        var cnt=0;
        for(var i=0; i<rows.length; i++){
          if((!(rows[i].id.indexOf('.')>0)) || rows[i].deleted){
            continue;
          }
          $scope.allSheetData[rows[i].id]=rows[i].doc;
          var tid=rows[i].id.substring(0, rows[i].id.indexOf('.'));
          if($scope.sheets.indexOf(tid)<0){
            $scope.sheets.push(tid);
          }
          if(!$scope.sheetName){
            $scope.sheetName=tid;
          }
          if($scope.sheetName!=tid){
            continue;
          }
          var rid=rows[i].id.substring(rows[i].id.indexOf('.')+1);
          $scope.sheetData[rid]=(rows[i].doc.formulares?rows[i].doc.formulares:(rows[i].doc.formatted?rows[i].doc.formatted:rows[i].doc.value));
          var maxCol=colToNum(rid);
          for(var j=$scope.sheetCols.length; j<=maxCol; j++){
            $scope.sheetCols.push(numToCol(j));
          }
          var maxRow=parseInt(rid.match(/[0-9]+/)[0]);
          for(var j=$scope.sheetRows.length; j<maxRow; j++){
            $scope.sheetRows.push(j+1);
          }
          if($scope.prevSince>0 && $scope.currentUpload!=rows[i].id){
            $scope.updatedCells.push(rid);
            cnt++;
          }
        }
        $scope.currentUpload=null;
        $scope.prevSince=data.last_seq;
        $scope.onFilter();
        $timeout(function(){
          for(var i=0; i<cnt; i++){
            $scope.updatedCells.shift();
          }
        }, 2500);
        if($scope.pollLimit == undefined || (($scope.pollLimit--) > 0)){ // enable testing
          pollForChanges();
        }
      }).error(function(data, status){
        if(status!=0){
          console.error(status+" - "+JSON.stringify(data));
          $scope.errorMessage="Error pulling data - refresh and try again ("+status+" - "+JSON.stringify(data)+")";
        }
        // other idea: poll also after a greace period - don't give up on one error (instead of F5-ing)
      });
  };
  function pollForUserName(){
    $http({method: "GET", url: "../../../../../_session?basic=true"})
      .success(function(data, status){
        $scope.userName=(data.userCtx && data.userCtx.name)?data.userCtx.name:"Webview";
      }).error(function(data, status){
        if(status!=0 && status!=401){ // 401=unauth (can be - in local sc)
          console.error(status+" - "+JSON.stringify(data));
          $scope.errorMessage="Error pulling user name - refresh and try again ("+status+" - "+JSON.stringify(data)+")";
        }
        $scope.userName="Webview";
      });
  };
  function doesContainAll(haystack, needle){
    haystack=haystack.toLowerCase();
    for(var i=0; i<needle.length; i++){
      if(haystack.indexOf(needle[i].toLowerCase())<0){
        return false;
      }
    }
    return true;
  };
  $scope.onFilter=function(){
    //console.log("onFilter");
    if($scope.filter){
      $scope.filteredRows=[];
      for(var i=0; i<$scope.sheetRows.length; i++){
        for(var j=0; j<$scope.sheetCols.length; j++){
          var addr=$scope.sheetCols[j]+$scope.sheetRows[i];
          if($scope.sheetData[addr] && doesContainAll((""+$scope.sheetData[addr]), $scope.filter.split(" "))){
            $scope.filteredRows.push($scope.sheetRows[i]);
            break;
          }
        }
      }
    }else{
      $scope.filteredRows=$scope.sheetRows;
    }
  };
  function clearEditElem(restoreValue){
    if($scope.editCell && $scope.isWebkit()){
      var editElem=$window.document.getElementsByName($scope.editCell)[0];
      while(editElem.childNodes.length>1){
        editElem.removeChild(editElem.childNodes[0]);
      }
      if(restoreValue){
        editElem.childNodes[0].nodeValue=restoreValue;
      }
    }
  };
  $scope.onChangePrompt=function(){
    if($scope.selectedCell){
      var oldValue=$scope.sheetData[$scope.selectedCell];
      var newValue=$window.prompt("Enter new value...", oldValue);
      if(newValue!=null && newValue!=oldValue){
        performUpload($scope.selectedCell, oldValue, newValue);
        $scope.sheetData[$scope.selectedCell]=newValue;
      }
    }
  };
  function confirmEdit(){
    if($scope.editCell && $scope.isWebkit()){
      //console.log("confirming edit");
      var oldValue=$scope.editBackup?$scope.editBackup:$scope.sheetData[$scope.editCell];
      var newValue=$window.document.getElementsByName($scope.editCell)[0].innerText;
      if(oldValue!=newValue){
        performUpload($scope.editCell, oldValue, newValue);
      }
      $scope.editBackup=null;
      // (always restore due to backup)
      $scope.sheetData[$scope.editCell]=newValue;
      clearEditElem();
      $scope.editCell=null;
    }
  };
  function performUpload(cell, oldValue, newValue){
    //console.log("Need to upload '"+cell+"': '"+oldValue+"' != '"+newValue+"' for user '"+$scope.userName+"'");
    $scope.currentUpload=$scope.sheetName+"."+cell;
    var upd={user: $scope.userName, type: "TEXT", value: (""+newValue)};
    if($scope.allSheetData[$scope.sheetName+"."+cell]){
      upd["_rev"]=$scope.allSheetData[$scope.sheetName+"."+cell]._rev;
    }
    $http({method: "PUT", url: "../../../../_design/cellupd/_update/cell/"+$scope.sheetName+"."+cell, data: JSON.stringify(upd)})
      .success(function(data, status){
      }).error(function(data, status){
        if(status!=0){
          console.error(status+" - "+JSON.stringify(data));
          $scope.errorMessage="Error updating - refresh and try updating again ("+status+" - "+JSON.stringify(data)+")";
        }
      });
  };
  function discardEdit(){
    if($scope.editCell && $scope.isWebkit()){
      //console.log("discarding edit");
      if($scope.editBackup){
        $scope.sheetData[$scope.editCell]=$scope.editBackup;
      }
      $scope.editBackup=null;
      clearEditElem($scope.sheetData[$scope.editCell]);
      $scope.editCell=null;
    }
  };
  $scope.onCellSelect=function(col, row, $event){
    //console.log("onCellSelect: "+(""+col+""+row));
    var newSel=(""+col+""+row);
    var dblClick=($scope.lastSelTimeStamp && ($event.timeStamp-$scope.lastSelTimeStamp)<250 && $scope.selectedCell==newSel);
    $scope.selectedCell=newSel;
    $scope.lastSelTimeStamp=$event.timeStamp;
    if($scope.editCell!=$scope.selectedCell){
      confirmEdit();
      if(dblClick && $scope.isWebkit()){
        $scope.editCell=$scope.selectedCell;
        // requeue
        $timeout(function(){
          //console.log("Inserting text");
          $window.document.getElementsByName($scope.editCell)[0].focus();
        }, 1);
      }else{
        $scope.editCell=null;
      }
    }
  };
  $scope.onEditCellBlur=function(col, row){
    //console.log("onEditCellBlur: "+(""+col+""+row));
    if($scope.editCell==(""+col+""+row)){
      confirmEdit();
    }
  };
  $scope.onScrollBottom=function(){
    var elem=$window.document.getElementsByName("A"+$scope.sheetRows[$scope.sheetRows.length-1])[0];
    smartScroll(elem);
  };
  function smartScroll(elem){
    if(elem.scrollIntoViewIfNeeded){
      elem.scrollIntoViewIfNeeded();
    }else{
      elem.scrollIntoView();
    }
  };
  function navigateToSel(){
    if($scope.timeoutRunning){
      return; // don't need two updates
    }
    $scope.timeoutRunning=true;
    $timeout(function(){
      $scope.timeoutRunning=false;
      var elem=$window.document.getElementsByName($scope.selectedCell)[0];
      smartScroll(elem);
    }, 500);
  };
  function focusTable(){
    // requeue
    $timeout(function(){
      $window.document.getElementsByTagName("table")[0].focus();
    }, 1);
  };
  $scope.onTableKeyDown=function($event){
    //console.log("onTableKeyDown: "+$event.keyCode);
    if($event.keyCode===27){
      // ESC (reset)
      discardEdit();
      $scope.editCell=null;
      $event.preventDefault();
      focusTable();
    }else if($event.keyCode===37 && (!$scope.editCell)){
      // left
      $scope.selectedCell=changeAddr($scope.selectedCell, -1, 0);
      navigateToSel();
      $event.preventDefault();
    }else if($event.keyCode===38 && (!$scope.editCell)){
      // up
      $scope.selectedCell=changeAddr($scope.selectedCell, 0, -1);
      navigateToSel();
      $event.preventDefault();
    }else if($event.keyCode===39 && (!$scope.editCell)){
      // right
      $scope.selectedCell=changeAddr($scope.selectedCell, 1, 0);
      navigateToSel();
      $event.preventDefault();
    }else if($event.keyCode===40 && (!$scope.editCell)){
      // down
      $scope.selectedCell=changeAddr($scope.selectedCell, 0, 1);
      navigateToSel();
      $event.preventDefault();
    }else if($event.keyCode===13){
      // ENTER (confirm edit / down)
      confirmEdit();
      $scope.editCell=null;
      $scope.selectedCell=changeAddr($scope.selectedCell, 0, 1);
      $event.preventDefault();
      focusTable();
    }else if($event.keyCode===9){
      // TAB (confirm edit / down)
      confirmEdit();
      $scope.editCell=null;
      $scope.selectedCell=changeAddr($scope.selectedCell, 1, 0);
      $event.preventDefault();
      focusTable();
    }else if($event.keyCode===8 && (!$scope.editCell)){
      // BACKSPACE (different handling: deactivate)
      $event.preventDefault();
    }
  };
  $scope.onTableKeyPress=function($event){
    //console.log("onTableKeyPress: "+$event.keyCode);
    if($event.keyCode===13 || $event.keyCode===9){
      // do nothing except confirm
      $event.preventDefault();
    }else if($scope.editCell!=$scope.selectedCell && $scope.isWebkit()){
      $scope.editBackup=$scope.sheetData[$scope.selectedCell];
      $scope.sheetData[$scope.selectedCell]="";
      $scope.editCell=$scope.selectedCell;
      $event.preventDefault();
      // requeue
      $timeout(function(){
        //console.log("Inserting text");
        $window.document.getElementsByName($scope.editCell)[0].focus();
        $window.document.execCommand("insertText", false, String.fromCharCode($event.keyCode));
      }, 1);
    }
  };
  $scope.addRow=function($event){
    //console.log("addRow");
    $scope.sheetRows.push($scope.sheetRows.length+1);
    // requeue
    $timeout(function(){
      smartScroll($event.target);
    }, 1);
  };
  $scope.addCol=function($event){
    //console.log("addCol");
    $scope.sheetCols.push(numToCol($scope.sheetCols.length));
    // requeue
    $timeout(function(){
      smartScroll($event.target);
    }, 1);
  };
  $scope.isSelected=function(col, row){
    return $scope.selectedCell==(""+col+""+row);
  };
  $scope.isEditable=function(col, row){
    return $scope.editCell==(""+col+""+row);
  };
  $scope.cellcontent=function(col, row){
    var res=$scope.sheetData[""+col+""+row];
    if(!res){
      res=" ";
    }
    return res;
  };
  $scope.cellstyle=function(col, row){
    var res= 
      (($scope.updatedCells.indexOf(""+col+""+row)>=0)?"changed ":"")+
      (($scope.isSelected(col, row))?"selected ":"");
    return res;
  };
  $scope.isWebkit=function(){
    return $window.navigator.userAgent.toLowerCase().indexOf("webkit")>=0;
  }
  if(!$scope.noPoll){ // enable testing
    pollForUserName();
    pollForChanges();
  }
  // finally
  focusTable();
});

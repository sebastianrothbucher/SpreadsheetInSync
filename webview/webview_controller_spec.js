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
describe("SheetController base suite", function() {
  var $httpBackend, $scope, $timeout, $window={}, createController;
  beforeEach(module('theplugin'));
  beforeEach(inject(function($injector) {
    $httpBackend = $injector.get('$httpBackend');
    $httpBackend.when('GET', '../../../../../_session?basic=true').respond({"userCtx": {"name": "homer"}});
    $httpBackend.when('GET', '../../../../_changes?since=0&feed=longpoll&include_docs=true').respond({"results":[
{"id":"Tabelle1.A1","doc":{"_id":"Tabelle1.A1","_rev":"1-6d659b35637ad17781920896bdec3686","type":"TEXT","value":"TODO-Liste 05/2014","user":""}},{"id":"Tabelle1.C4","doc":{"_id":"Tabelle1.C4","_rev":"3-12dbf56a04a6f940e9886fd0a7730c8b","type":"VALUE","value":41805,"formatted":"15.06.14","user":""}}], "last_seq": 4});
    var $rootScope = $injector.get('$rootScope');
    $scope=$rootScope.$new();
    $scope.noPoll=true;
    $scope.pollLimit=0;
    $timeout=function(){};
    var $controller = $injector.get('$controller');
    createController = function() {
      return $controller('SheetController', {'$scope' : $scope, '$window': $window, '$timeout': $timeout});
    };
  }));
  it("contains cell A1 to C4", function() {
    $httpBackend.expectGET('../../../../../_session?basic=true');
    $httpBackend.expectGET('../../../../_changes?since=0&feed=longpoll&include_docs=true');
    $scope.noPoll=false;
    createController();
    $httpBackend.flush();
    expect($scope.sheetCols.length).toEqual(3);
    expect($scope.sheetRows.length).toEqual(4);
    expect($scope.filteredRows.length).toEqual(4);
  });
  it("has A1 Headline TODO-Liste 05/2014", function() {
    $httpBackend.expectGET('../../../../../_session?basic=true');
    $httpBackend.expectGET('../../../../_changes?since=0&feed=longpoll&include_docs=true');
    $scope.noPoll=false;
    createController();
    $httpBackend.flush();
    expect($scope.sheetData["A1"]).toEqual("TODO-Liste 05/2014");
  });
  it("has C4 content formatted", function() {
    $httpBackend.expectGET('../../../../../_session?basic=true');
    $httpBackend.expectGET('../../../../_changes?since=0&feed=longpoll&include_docs=true');
    $scope.noPoll=false;
    createController();
    $httpBackend.flush();
    expect($scope.sheetData["C4"]).toEqual("15.06.14");
  });
  it("updates and shows updates", function() {
    $httpBackend.when('GET', '../../../../_changes?since=4&feed=longpoll&include_docs=true').respond({"results":[
{"id":"Tabelle1.C3","doc":{"_id":"Tabelle1.C3","_rev":"3-12dbf56a04a6f940e9886fd0a7730c8b","type":"TEXT","value":"test","user":""}}], "last_seq": 5});
    $httpBackend.expectGET('../../../../../_session?basic=true');
    $httpBackend.expectGET('../../../../_changes?since=0&feed=longpoll&include_docs=true');
    $httpBackend.expectGET('../../../../_changes?since=4&feed=longpoll&include_docs=true');
    $scope.noPoll=false;
    $scope.pollLimit=1;
    var cb;
    $timeout=function(cb_, to_){
      cb=cb_;
    };
    createController();
    $httpBackend.flush();
    expect($scope.sheetData["C3"]).toEqual("test");
    expect($scope.updatedCells.length).toEqual(1);
    expect($scope.updatedCells[0]).toEqual("C3");
    cb();
    expect($scope.updatedCells.length).toEqual(0);
  });
  it("Performs filter based on any cell with filter as part", function(){
    createController();
    $scope.sheetCols=["A", "B", "C"];
    $scope.sheetRows=[1, 2, 3, 4, 5, 6, 7, 8, 9];
    $scope.sheetData={"A1": "blup bla hey", "A2": "bla", "A3": "blup", "B3": "hey bla"};
    $scope.filter="bla hey";
    $scope.onFilter();
    expect($scope.filteredRows.length).toBe(2);
    expect($scope.filteredRows.indexOf(1)>=0).toBeTruthy();
    expect($scope.filteredRows.indexOf(2)>=0).toBeFalsy();
    expect($scope.filteredRows.indexOf(3)>=0).toBeTruthy();
    $scope.filter=""; // (clear again)
    $scope.onFilter();
    expect($scope.filteredRows.length).toBe(9);
    expect($scope.filteredRows.indexOf(1)>=0).toBeTruthy();
    expect($scope.filteredRows.indexOf(2)>=0).toBeTruthy();
    expect($scope.filteredRows.indexOf(3)>=0).toBeTruthy();
  });
  it("Styles cells", function(){
    createController();
    $scope.selectedCell="B2";
    $scope.updatedCells=["D3"];
    expect($scope.cellstyle("B", 2)).toMatch(".*selected.*");
    expect($scope.cellstyle("B", 3)).not.toMatch(".*selected.*");
    expect($scope.cellstyle("D", 3)).toMatch(".*changed.*");
    expect($scope.cellstyle("D", 4)).not.toMatch(".*changed.*")
  });
  describe("Navigating", function(){
    var focused, scrolled;
    beforeEach(function(){
      focused=[];
      scrolled=[];
      $window.navigator={
        "userAgent": ""
      };
      $timeout=function(cb_, to_){
        cb_(); // just do it
      };
      $window.document = {
        "getElementsByName": function(n_){
          return [
            {
              "focus": function(){
                focused.push(n_);
              },
              "scrollIntoViewIfNeeded": function(){
                scrolled.push(n_);
              }
            }
          ];
        },
        "getElementsByTagName": function(n_){
          return [
            {
              "focus": function(){}
            }
          ];
        }
      };
    });
    it("Changes cell based on click", function(){
      createController();
      $scope.selectedCell="D5";
      var ev={
        "timeStamp": 0,
        "preventDefault": function(){}
      };
      $scope.onCellSelect("B", 2, ev);
      expect($scope.selectedCell).toBe("B2");
    });
    it("Changes cell based on cursor keys considering filter", function(){
      createController();
      $scope.selectedCell="D5";
      $scope.sheetCols=["A", "B", "C", "D", "E"];
      $scope.sheetRows=[1, 2, 3, 4, 5, 6, 7];
      $scope.filteredRows=$scope.sheetRows.filter(function(r){return r!=4});
      var ev={
        "preventDefault": function(){}
      };
      spyOn(ev, 'preventDefault');
      ev.preventDefault.calls.reset();
      ev.keyCode=37; // left arrow
      $scope.onTableKeyDown(ev);
      expect(ev.preventDefault).toHaveBeenCalled();
      expect($scope.selectedCell).toBe("C5");
      expect(scrolled.length).toBe(1);
      expect(scrolled[0]).toBe("C5");
      ev.preventDefault.calls.reset();
      ev.keyCode=38; // up arrow
      $scope.onTableKeyDown(ev);
      expect(ev.preventDefault).toHaveBeenCalled();
      expect($scope.selectedCell).toBe("C3");
      expect(scrolled.length).toBe(2);
      expect(scrolled[1]).toBe("C3");
      ev.preventDefault.calls.reset();
      ev.keyCode=39; // right arrow
      $scope.onTableKeyDown(ev);
      expect(ev.preventDefault).toHaveBeenCalled();
      expect($scope.selectedCell).toBe("D3");
      expect(scrolled.length).toBe(3);
      expect(scrolled[2]).toBe("D3");
      ev.preventDefault.calls.reset();
      ev.keyCode=40; // down arrow
      $scope.onTableKeyDown(ev);
      expect(ev.preventDefault).toHaveBeenCalled();
      expect($scope.selectedCell).toBe("D5");
      expect(scrolled.length).toBe(4);
      expect(scrolled[3]).toBe("D5");
    });
    it("Changes cell based on ENTER and TAB", function(){
      createController();
      $scope.selectedCell="D5";
      $scope.sheetCols=["A", "B", "C", "D", "E"];
      $scope.sheetRows=[1, 2, 3, 4, 5, 6, 7];
      $scope.filteredRows=$scope.sheetRows;
      var ev={
        "preventDefault": function(){}
      };
      spyOn(ev, 'preventDefault');
      ev.preventDefault.calls.reset();
      ev.keyCode=13; // ENTER
      $scope.onTableKeyDown(ev);
      expect(ev.preventDefault).toHaveBeenCalled();
      expect($scope.selectedCell).toBe("D6");
      ev.preventDefault.calls.reset();
      ev.keyCode=9; // TAB
      $scope.onTableKeyDown(ev);
      expect(ev.preventDefault).toHaveBeenCalled();
      expect($scope.selectedCell).toBe("E6");
      ev.preventDefault.calls.reset();
    });
  });
  describe("Base editing", function(){
    it("Updates on prompt", function(){
      $httpBackend.when('PUT', '../../../../_design/cellupd/_update/cell/Tabelle1.B3', function(str){return JSON.parse(str).value=="new fragrance"}).respond({"ok": true});
      $httpBackend.expectPUT('../../../../_design/cellupd/_update/cell/Tabelle1.B3');
      createController();
      $scope.allSheetData['Tabelle1.B3']={"_rev": "1-123"};
      $scope.sheetName="Tabelle1"
      $scope.sheetData['B3']="old spice";
      $scope.selectedCell="B3";
      $window.prompt=function(){return "new fragrance";};
      spyOn($window, "prompt").and.callThrough();
      $scope.onChangePrompt();
      $httpBackend.flush();
      expect($scope.sheetData['B3']).toBe("new fragrance");
      expect($window.prompt).toHaveBeenCalledWith("Enter new value...", "old spice");
    });
    it("Does nothing on same value", function(){
      createController();
      $scope.sheetName="Tabelle1"
      $scope.sheetData['B3']="old spice";
      $scope.selectedCell="B3";
      $window.prompt=function(){return "old spice";};
      spyOn($window, "prompt").and.callThrough();
      $scope.onChangePrompt();
      $httpBackend.verifyNoOutstandingExpectation();
    });
    it("Does nothing on cancel", function(){
      createController();
      $scope.sheetName="Tabelle1"
      $scope.sheetData['B3']="old spice";
      $scope.selectedCell="B3";
      $window.prompt=function(){return null;};
      spyOn($window, "prompt").and.callThrough();
      $scope.onChangePrompt();
      $httpBackend.verifyNoOutstandingExpectation();
    });
  });
  describe("Webkit inline editing", function(){
    var focused = false;
    var cmds=[];
    var params=[];
    beforeEach(function(){
      $window.navigator={
        "userAgent": "some webkit stuff"
      };
      $timeout=function(cb_, to_){
        cb_(); // just do it
      };
      $window.document = {
        "getElementsByName": function(n_){
          if("D5"==n_){
            return [
              {
                "focus": function(){
                  focused=true;
                }, 
                "childNodes": [
                  {
                    "nodeValue": ""
                  }
                ],
                "innerText": "new value"
              }
            ];
          }
        },
        "getElementsByTagName": function(n_){
          return [
            {
              "focus": function(){}
            }
          ];
        },
        "execCommand": function(cmd_, b_, param_){
          cmds.push(cmd_);
          params.push(param_);
        }
      };
    });
    it("Starts editing on type", function(){
      createController();
      $scope.selectedCell="D5";
      $scope.sheetData={"D5": "hey"};
      var ev={
        "keyCode": 97, // 'a'
        "preventDefault": function(){}
      };
      spyOn(ev, 'preventDefault');
      $scope.onTableKeyPress(ev);
      expect($scope.editCell).toBe("D5");
      expect(focused).toBeTruthy();
      expect(cmds.length).toBe(1);
      expect(params.length).toBe(1);
      expect(params[0]).toBe('a');
      expect(ev.preventDefault).toHaveBeenCalled();
    });
    it("Starts editing double-click", function(){
      createController();
      var ev={};
      ev.timeStamp=1436363001;
      $scope.onCellSelect("D", 5, ev);
      expect($scope.selectedCell).toBe("D5");
      expect($scope.editCell).toBeNull();
      ev.timeStamp=1436363101;
      $scope.onCellSelect("D", 5, ev);
      expect($scope.selectedCell).toBe("D5");
      expect($scope.editCell).toBe("D5");
      expect(focused).toBeTruthy();
    });
    it("Cancels editing on escape", function(){
      createController();
      $scope.editCell="D5";
      $scope.editBackup="old value";
      var ev={
        "keyCode": 27, // ESC
        "preventDefault": function(){}
      };
      spyOn(ev, 'preventDefault');
      $scope.onTableKeyDown(ev);
      expect(ev.preventDefault).toHaveBeenCalled();
      expect($scope.sheetData['D5']).toBe("old value");
      expect($scope.editCell).toBeNull();
    });
    it("Confirms editing on ENTER", function(){
      $httpBackend.when('PUT', '../../../../_design/cellupd/_update/cell/Tabelle1.D5', function(str){return JSON.parse(str).value=="new value"}).respond({"ok": true});
      $httpBackend.expectPUT('../../../../_design/cellupd/_update/cell/Tabelle1.D5');
      createController();
      $scope.selectedCell="D5";
      $scope.editCell="D5";
      $scope.sheetName="Tabelle1";
      $scope.sheetCols=["A", "B", "C", "D", "E"];
      $scope.sheetRows=[1, 2, 3, 4, 5, 6];
      $scope.filteredRows=$scope.sheetRows;
      var ev={
        "keyCode": 13, // ENTER
        "preventDefault": function(){}
      };
      spyOn(ev, 'preventDefault');
      $scope.onTableKeyDown(ev);
      expect(ev.preventDefault).toHaveBeenCalled();
      $httpBackend.flush();
      expect($scope.sheetData['D5']).toBe("new value");
      expect($scope.selectedCell).toBe("D6");
      expect($scope.editCell).toBeNull;
    });
    it("Confirms editing on TAB", function(){
      $httpBackend.when('PUT', '../../../../_design/cellupd/_update/cell/Tabelle1.D5', function(str){return JSON.parse(str).value=="new value"}).respond({"ok": true});
      $httpBackend.expectPUT('../../../../_design/cellupd/_update/cell/Tabelle1.D5');
      createController();
      $scope.selectedCell="D5";
      $scope.editCell="D5";
      $scope.sheetName="Tabelle1";
      $scope.sheetCols=["A", "B", "C", "D", "E"];
      $scope.sheetRows=[1, 2, 3, 4, 5, 6];
      $scope.filteredRows=$scope.sheetRows;
      var ev={
        "keyCode": 9, // TAB
        "preventDefault": function(){}
      };
      spyOn(ev, 'preventDefault');
      $scope.onTableKeyDown(ev);
      expect(ev.preventDefault).toHaveBeenCalled();
      $httpBackend.flush();
      expect($scope.sheetData['D5']).toBe("new value");
      expect($scope.selectedCell).toBe("E5");
      expect($scope.editCell).toBeNull;
    });
    it("Confirms editing on selecting something else", function(){
      $httpBackend.when('PUT', '../../../../_design/cellupd/_update/cell/Tabelle1.D5', function(str){return JSON.parse(str).value=="new value"}).respond({"ok": true});
      $httpBackend.expectPUT('../../../../_design/cellupd/_update/cell/Tabelle1.D5');
      createController();
      $scope.selectedCell="D5";
      $scope.editCell="D5";
      $scope.sheetName="Tabelle1";
      $scope.sheetCols=["A", "B", "C", "D", "E"];
      $scope.sheetRows=[1, 2, 3, 4, 5, 6];
      var ev={
        "timeStamp": 0,
        "preventDefault": function(){}
      };
      $scope.onCellSelect("B", 2, ev);
      $httpBackend.flush();
      expect($scope.sheetData['D5']).toBe("new value");
      expect($scope.selectedCell).toBe("B2");
      expect($scope.editCell).toBeNull;
    });
    it("Does nothing on same cell value", function(){
      createController();
      $scope.selectedCell="D5";
      $scope.editCell="D5";
      $scope.editBackup="new value"; // (already)
      $scope.sheetName="Tabelle1";
      $scope.sheetCols=["A", "B", "C", "D", "E"];
      $scope.sheetRows=[1, 2, 3, 4, 5, 6];
      var ev={
        "timeStamp": 0,
        "preventDefault": function(){}
      };
      $scope.onCellSelect("B", 2, ev);
      expect($scope.sheetData['D5']).toBe("new value");
      expect($scope.selectedCell).toBe("B2");
      expect($scope.editCell).toBeNull;
      $httpBackend.verifyNoOutstandingExpectation();
    });
    it("Does nothing on non-webkit", function(){
      $window.navigator={
        "userAgent": "some other stuff"
      };
      createController();
      $scope.selectedCell="D5";
      var ev={
        "keyCode": 97, // 'a'
        "preventDefault": function(){}
      };
      spyOn(ev, 'preventDefault');
      $scope.onTableKeyPress(ev);
      expect($scope.editCell).toBeUndefined();
    });
  });
}); 

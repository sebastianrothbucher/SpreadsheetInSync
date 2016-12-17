#!bash
## Licensed under the Apache License, Version 2.0 (the "License"); you may not
## use this file except in compliance with the License. You may obtain a copy of
## the License at
##
##   http://www.apache.org/licenses/LICENSE-2.0
##
## Unless required by applicable law or agreed to in writing, software
## distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
## WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
## License for the specific language governing permissions and limitations under
## the License.
if [ ! $1 ]; then echo "No target DB URL"; exit 1; fi
curl -X PUT $1/angular -d '{}'
ANGVERSION=$(curl -X GET $1/angular | node -e 'var dt=""; process.stdin.on("data", function (data){dt+=data;}).on("end", function(){console.log(JSON.parse(dt)["_rev"]);});')
curl -X PUT $1/angular/angular.min.js?rev=$ANGVERSION -H 'Content-type: application/javascript' --data-binary @angular/angular.min.js
node -e 'var fs=require("fs"); fs.readFile("showfkt.json", "utf-8", function(err, data){var dt=JSON.parse(data); delete dt["_rev"]; console.log(JSON.stringify(dt));});' | curl -X PUT $1/_design/showfkt -d @-
node -e 'var fs=require("fs"); fs.readFile("webview.htm", "utf-8", function(err, data){console.log(JSON.stringify({"html": data, "type": "text/html"}));});' | curl -X PUT $1/webview -d @-
node -e 'var fs=require("fs"); fs.readFile("webview_controller.js", "utf-8", function(err, data){console.log(JSON.stringify({"html": data, "type": "text/javascript"}));});' | curl -X PUT $1/webview_controller -d @-


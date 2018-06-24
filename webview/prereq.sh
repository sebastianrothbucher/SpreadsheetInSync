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

if [ ! -d angular ]; then
	mkdir angular
fi
if [ ! -f angular/angular.js ]; then
	echo "need angular/angular.js"
	curl -X GET -o angular/angular.js https://code.angularjs.org/1.2.15/angular.js
fi
if [ ! -f angular/angular.min.js ]; then
	echo "need angular/angular.min.js"
	curl -X GET -o angular/angular.min.js https://code.angularjs.org/1.2.15/angular.min.js
fi
if [ ! -f angular/angular-mocks.js ]; then
	echo "need angular/angular-mocks.js"
	curl -X GET -o angular/angular-mocks.js https://code.angularjs.org/1.2.15/angular-mocks.js
fi
npm install


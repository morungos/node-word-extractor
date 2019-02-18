### 0.2.2 / 23rd January 2019

 * Fixed [the bad dependency on event-stream](https://github.com/dominictarr/event-stream/issues/116)


### 0.3.0 / 18th February 2019

 * Re-fixed the bad loop in the OLE code. See #15, #18
 * A few errors previously rejected as strings, they're now errors
 * Updated dependencies to safe versions. See #20


### 0.2.1 / 21st January 2019

 * Added a new getEndnotes method. See #16
 * Fixed a bad loop in the OLE code


### 0.2.0 / 31st October 3018

 * Removed coffeescript and mocha, now using jest and plain ES6
 * Removed partial work on .docx (for now)


### 0.1.4 / 25th March 2017

 * Fixed a documentation issue. `extract` returns a Promise. See #6
 * Corrected table cell delimiters to be tabs. See #9
 * Fixed an issue where replacements weren't being applied right. 


### 0.1.3 / 6th July 2016

 * Added the missing `lib` folder
 * Added a missing dependency to `package.json`


### 0.1.1 / 17th January 2016

 * Fixed a bug with text boundary calculations
 * Added endpoints `getHeaders`, `getFootnotes`, `getAnnotations`


### 0.1.0 / 14th January 2016

 * Initial release to npm

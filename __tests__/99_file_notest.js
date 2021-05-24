/**
 * A handy tool we can use to debug tests, by applying to a given file.
 */

const path = require('path');
const WordExtractor = require('../lib/word');

const extractor = new WordExtractor();
extractor.extract(path.resolve(__dirname, process.argv[2]))
  .then((d) => console.log(d))
  .catch((e) => console.error(e));

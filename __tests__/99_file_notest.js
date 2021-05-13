/**
 * A handy tool we can use to debug tests, by applying to a given file.
 */

const path = require('path');

const WordExtractor = require('../lib/word');

const extractor = new WordExtractor();
extractor.extract(path.resolve(__dirname, "data/test01.doc"));

## word-extractor

![test workflow](https://github.com/morungos/node-word-extractor/actions/workflows/main.yml/badge.svg)

Read data from a Word document (.doc or .docx) using Node.js

### Why use this module?

There are a fair number of npm components which can extract text from Word .doc
files, but they all appear to require some external helper program, and involve
either spawning a process or communicating with a persistent one. That raises
the installation and deployment burden as well as the runtime one.

This module is intended to provide a much faster way of reading the text from a
Word file, without leaving the Node.js environment.

This means you do not need to install Word, Office, or anything else, and the
module will work on all platforms, without any native binary code requirements.

As of version 1.0, this module supports both traditional, OLE-based, Word files (usually .doc), 
and modern, Open Office-style, ECMA-376 Word files (usually .docx). It can be 
used both with files and with file contents in a Node.js Buffer.

### How do I install this module?

```bash=
yarn add word-extractor

# Or using npm... 
npm install word-extractor
```

### How do I use this module?

```
const WordExtractor = require("word-extractor"); 
const extractor = new WordExtractor();
const extracted = extractor.extract("file.doc");

extracted.then(function(doc) { console.log(doc.getBody()); });
```

The object returned from the `extract()` method is a promise that resolves to a
document object, which then provides several views onto different parts of the
document contents.

### Methods

`WordExtractor#extract(<filename>  | <Buffer>)`

Main method to open a Word file and retrieve the data. Returns a promise which
resolves to a `Document`. If a Buffer is passed instead of a filename, then
the buffer is used directly, instad of reading a disk from the file system.

`Document#getBody()`

Retrieves the content text from a Word document. This will handle UNICODE
characters correctly, so if there are accented or non-Latin-1 characters
present in the document, they'll show as is in the returned string.

`Document#getFootnotes()`

Retrieves the footnote text from a Word document. This will handle UNICODE
characters correctly, so if there are accented or non-Latin-1 characters
present in the document, they'll show as is in the returned string.

`Document#getEndnotes()`

Retrieves the endnote text from a Word document. This will handle UNICODE
characters correctly, so if there are accented or non-Latin-1 characters
present in the document, they'll show as is in the returned string.

`Document#getHeaders()`

Retrieves the header and footer text from a Word document. This will handle
UNICODE characters correctly, so if there are accented or non-Latin-1
characters present in the document, they'll show as is in the returned string.

`Document#getAnnotations()`

Retrieves the comment bubble text from a Word document. This will handle
UNICODE characters correctly, so if there are accented or non-Latin-1
characters present in the document, they'll show as is in the returned string.

### License

Copyright (c) 2016-2021. Stuart Watt.

Licensed under the MIT License.

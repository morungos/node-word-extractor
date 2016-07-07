### word-extractor

Read data from a Word document using node.js


#### Why use this module?

There are a fair number of npm components which can extract text from Word .doc files, but they all appear to require some external helper program, and involve either spawning a process or communicating with a persistent one. That raises the installation and deployment burden as well as the runtime one.

This module is intended to provide a much faster way of reading the text from a Word file, without leaving the node.js environment.


#### How do I use this module?

    var WordExtractor = require("word-extractor");
    var extractor = new WordExtractor();
    var doc = extractor.extract("file.doc")
    var body = doc.getBody();

The object returned from the `extract()` method is a document object, and provides several views onto different parts of the document contents.


#### Methods

`WordExtractor#extract(file)`

Main method to open a Word file and retrieve the data. Returns a `Document`.

`Document#getBody()`

Retrieves the content text from a Word document. This will handle UNICODE characters correctly, so if there are accented or non-Latin-1 characters present in the document, they'll show as is in the returned string.

`Document#getFootnotes()`

Retrieves the footnote text from a Word document. This will handle UNICODE characters correctly, so if there are accented or non-Latin-1 characters present in the document, they'll show as is in the returned string.

`Document#getHeaders()`

Retrieves the header and footer text from a Word document. This will handle UNICODE characters correctly, so if there are accented or non-Latin-1 characters present in the document, they'll show as is in the returned string.

`Document#getAnnotations()`

Retrieves the comment bubble text from a Word document. This will handle UNICODE characters correctly, so if there are accented or non-Latin-1 characters present in the document, they'll show as is in the returned string.


#### License

Copyright (c) 2016. Stuart Watt.

Licensed under the MIT License.

/**
 * @module filters
 * 
 * @description
 * Exports several functions that implement various methods for translating
 * characters into Unicode, and cleaning up some of the remaining residues from
 * Word's odd internal marker character usage.
 */

const replaceTable = [];
replaceTable[0x0002] = '\x00';
replaceTable[0x0005] = '\x00';
replaceTable[0x0007] = "\t";
replaceTable[0x0008] = '\x00';
replaceTable[0x000A] = "\n";
replaceTable[0x000D] = "\n";
replaceTable[0x001E] = "\u2011";

const binaryToUnicodeTable = [];
binaryToUnicodeTable[0x0082] = "\u201a";
binaryToUnicodeTable[0x0083] = "\u0192";
binaryToUnicodeTable[0x0084] = "\u201e";
binaryToUnicodeTable[0x0085] = "\u2026";
binaryToUnicodeTable[0x0086] = "\u2020";
binaryToUnicodeTable[0x0087] = "\u2021";
binaryToUnicodeTable[0x0088] = "\u02C6";
binaryToUnicodeTable[0x0089] = "\u2030";
binaryToUnicodeTable[0x008a] = "\u0160";
binaryToUnicodeTable[0x008b] = "\u2039";
binaryToUnicodeTable[0x008c] = "\u0152";
binaryToUnicodeTable[0x008e] = "\u017D";
binaryToUnicodeTable[0x0091] = "\u2018";
binaryToUnicodeTable[0x0092] = "\u2019";
binaryToUnicodeTable[0x0093] = "\u201C";
binaryToUnicodeTable[0x0094] = "\u201D";
binaryToUnicodeTable[0x0095] = "\u2022";
binaryToUnicodeTable[0x0096] = "\u2013";
binaryToUnicodeTable[0x0097] = "\u2014";
binaryToUnicodeTable[0x0098] = "\u02DC";
binaryToUnicodeTable[0x0099] = "\u2122";
binaryToUnicodeTable[0x009a] = "\u0161";
binaryToUnicodeTable[0x009b] = "\u203A";
binaryToUnicodeTable[0x009c] = "\u0153";
binaryToUnicodeTable[0x009e] = "\u017E";
binaryToUnicodeTable[0x009f] = "\u0178";

/**
* Replaces Word characters with an appropriate substitution.
* @param {*} string an input string
* @returns a modified string
*/
const replace = (string) => {
  return string.replace(/([\x02\x05\x07\x08\x0a\x0d])/g, (match) => replaceTable[match.charCodeAt(0)]);
};

const binaryToUnicode = (string) => {
  return string.replace(/([\x80-\x9f])/g, (match) => binaryToUnicodeTable[match.charCodeAt(0)]);
};

const fieldReplacer = /(?:\x13(?:[^\x13\x14\x15]*\x14)?([^\x13\x14\x15]*)\x15)/g;

const clean = (string) => {

  // Fields can be nested, which makes this awkward. We use a strict non-nesting model
  // and repeat until we find no substitutions. This is because a second match might
  // start before an earlier one, due to our replacements.

  let called = true;
  while (called) {
    called = false;
    string = string.replace(fieldReplacer, (match, p1) => { called = true; return p1; });
  }

  return string
    .replace(/[\x00-\x07]/g, '');
};

const filterTable = [];
filterTable[0x2002] = " ";
filterTable[0x2003] = " ";
filterTable[0x2012] = "-";
filterTable[0x2013] = "-";
filterTable[0x2014] = "-";
filterTable[0x2018] = "'";
filterTable[0x2019] = "'";
filterTable[0x201c] = "\"";
filterTable[0x201d] = "\"";

const filter = (string) => {
  return string
    .replace(/[\u2002\u2003\u2012\u2013\u2014\u2018\u2019\u201c\u201d]/g, (match) => filterTable[match.charCodeAt(0)]);
};

module.exports = {
  clean: clean,
  binaryToUnicode: binaryToUnicode,
  replace: replace,
  filter: filter
};

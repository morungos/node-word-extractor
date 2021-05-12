const replacements = [];
replacements[0x0002] = '\x00';
replacements[0x0005] = '\x00';
replacements[0x0007] = "\t";
replacements[0x0008] = '\x00';
replacements[0x000A] = "\n";
replacements[0x000D] = "\n";

module.exports.replace = (string) => {
  return string.replace(/([\x02\x05\x07\x08\x0a\x0d])/g, (match) => replacements[match.charCodeAt(0)]);
};

const binaryToUnicode = [];
binaryToUnicode[0x0082] = "\u201a";
binaryToUnicode[0x0083] = "\u0192";
binaryToUnicode[0x0084] = "\u201e";
binaryToUnicode[0x0085] = "\u2026";
binaryToUnicode[0x0086] = "\u2020";
binaryToUnicode[0x0087] = "\u2021";
binaryToUnicode[0x0088] = "\u02C6";
binaryToUnicode[0x0089] = "\u2030";
binaryToUnicode[0x008a] = "\u0160";
binaryToUnicode[0x008b] = "\u2039";
binaryToUnicode[0x008c] = "\u0152";
binaryToUnicode[0x008e] = "\u017D";
binaryToUnicode[0x0091] = "\u2018";
binaryToUnicode[0x0092] = "\u2019";
binaryToUnicode[0x0093] = "\u201C";
binaryToUnicode[0x0094] = "\u201D";
binaryToUnicode[0x0095] = "\u2022";
binaryToUnicode[0x0096] = "\u2013";
binaryToUnicode[0x0097] = "\u2014";
binaryToUnicode[0x0098] = "\u02DC";
binaryToUnicode[0x0099] = "\u2122";
binaryToUnicode[0x009a] = "\u0161";
binaryToUnicode[0x009b] = "\u203A";
binaryToUnicode[0x009c] = "\u0153";
binaryToUnicode[0x009e] = "\u017E";
binaryToUnicode[0x009f] = "\u0178";

module.exports.binaryToUnicode = (string) => {
  return string.replace(/([\x80-\x9f])/g, (match) => binaryToUnicode[match.charCodeAt(0)]);
};

module.exports.clean = (string) => {
  return string
    .replace(/(?:\x13(?:[^\x14]*\x14)?([^\x15]*)\x15)/g, (match, p1) => p1)
    .replace(/\x00/g, '');
};

// const replacer = function(match, ...rest) {
//   if (match.length === 1) {
//     const replaced = filters[match.charCodeAt(0)];
//     if (replaced === 0) {
//       return "";
//     } else {
//       return replaced;
//     }
//   } else if (rest.length === 2) {
//     return "";
//   } else if (rest.length === 3) {
//     return rest[0];
//   }
// };

// const matcher = /(?:[\x02\x05\x07\x08\x0a\x0d\x80-\x9f\u2018\u2019\u201c\u201d\u2002\u2003\u2012\u2013\u2014])/g;

// const fieldFilter = /(?:\x13(?:[^\x14]*\x14)?([^\x15]*)\x15)/g;

// const filter = (text) => {
//   // const explore = text.replace(/(?:[\x00-\x1f\x7f-\uffff])/g, (match) => {
//   //   const code = match.charCodeAt(0).toString(16);
//   //   return '\\u' + '0000'.substring(0, 4 - code.length) + code;
//   // });
//   // if (explore !== text) 
//   //   console.log("result:", explore);
//   return text.replace(matcher, replacer).replace(fieldFilter, (match, p1) => p1);
// };

const filters = [];
filters[0x0002] = 0;
filters[0x0005] = 0;
filters[0x0008] = 0;
filters[0x2018] = "'";
filters[0x2019] = "'";
filters[0x201C] = "\"";
filters[0x201D] = "\"";
filters[0x0007] = "\t";
filters[0x000D] = "\n";
filters[0x2002] = " ";
filters[0x2003] = " ";
filters[0x2012] = "-";
filters[0x2013] = "-";
filters[0x2014] = "-";
filters[0x000A] = "\n";
filters[0x000D] = "\n";

const replacer = function(match, ...rest) {
  if (match.length === 1) {
    const replaced = filters[match.charCodeAt(0)];
    if (replaced === 0) {
      return "";
    } else {
      return replaced;
    }
  } else if (rest.length === 2) {
    return "";
  } else if (rest.length === 3) {
    return rest[0];
  }
};

const matcher = /(?:[\x02\x05\x07\x08\x0a\x0d\u2018\u2019\u201c\u201d\u2002\u2003\u2012\u2013\u2014])/g;

const fieldFilter = /(?:\x13(?:[^\x14]*\x14)?([^\x15]*)\x15)/g;

const filter = (text) => {
  return text.replace(matcher, replacer).replace(fieldFilter, (match, p1) => p1);
};


module.exports = {
  filters: filters,
  replacer: replacer,
  matcher: matcher,
  filter: filter
};

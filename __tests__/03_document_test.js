const Document = require('../lib/document');

describe('Document', () => {

  it('should instantiate successfully', () => {
    const document = new Document();
    expect(document).toBeInstanceOf(Document);
  });

  it('should read the body', () => {
    const document = new Document();
    document._body = "This is the body";
    expect(document.getBody()).toBe("This is the body");
  });

  it('should read the footnotes', () => {
    const document = new Document();
    document._footnotes = "This is the footnotes";
    expect(document.getFootnotes()).toBe("This is the footnotes");
  });

  it('should read the endnotes', () => {
    const document = new Document();
    document._endnotes = "This is the endnotes";
    expect(document.getEndnotes()).toBe("This is the endnotes");
  });

  it('should read the annotations', () => {
    const document = new Document();
    document._annotations = "This is the annotations";
    expect(document.getAnnotations()).toBe("This is the annotations");
  });

  it('should read the headers', () => {
    const document = new Document();
    document._headers = "This is the headers";
    expect(document.getHeaders()).toBe("This is the headers");
  });

  it('should read the headers and footers', () => {
    const document = new Document();
    document._headers = "This is the headers\n";
    document._footers = "This is the footers\n";
    expect(document.getHeaders()).toBe("This is the headers\nThis is the footers\n");
  });

  it('should selectively read the headers', () => {
    const document = new Document();
    document._headers = "This is the headers\n";
    document._footers = "This is the footers\n";
    expect(document.getHeaders({includeFooters: false})).toBe("This is the headers\n");
  });

  it('should read the footers', () => {
    const document = new Document();
    document._headers = "This is the headers\n";
    document._footers = "This is the footers\n";
    expect(document.getFooters()).toBe("This is the footers\n");
  });

});
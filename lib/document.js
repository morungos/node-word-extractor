/**
 * @module document
 * 
 * @description
 * Implements the main document returned when a Word file has been extracted. This exposes
 * methods that allow the body, annotations, headers, footnotes, and endnotes, to be 
 * read and used.
 * 
 * @author
 * Stuart Watt <stuart@morungos.com>
 */

const { filter } = require('./filters');

/**
 * @class
 * Returned from all extractors, this class provides accessors to read the different
 * parts of a Word document. This also allows some options to be passed to the accessors,
 * so you can control some character conversion and filtering, as described in the methods
 * below.
 */
class Document {
  
  constructor() {
    this._body = "";
    this._footnotes = "";
    this._endnotes = "";
    this._headers = "";
    this._footers = "";
    this._annotations = "";
  }

  /**
   * Accessor to read the main body part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @returns a string, containing the Word file body
   */
  getBody(options) {
    options = options || {};
    const value = this._body;
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the footnotes part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @returns a string, containing the Word file footnotes
   */
  getFootnotes(options) {
    options = options || {};
    const value = this._footnotes;
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the endnotes part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @returns a string, containing the Word file endnotes
   */
  getEndnotes(options) {
    options = options || {};
    const value = this._endnotes;
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the headers part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @param {boolean} options.includeFooters - if true (the default), returns headers and footers 
   *   as a single string
   * @returns a string, containing the Word file headers
   */
  getHeaders(options) {
    options = options || {};
    const value = this._headers + ((options.includeFooters == false) ? "" : this._footers);
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the footers part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @returns a string, containing the Word file footers
   */
  getFooters(options) {
    options = options || {};
    const value = this._footers;
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the annotations part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @returns a string, containing the Word file annotations
   */
  getAnnotations(options) {
    options = options || {};
    const value = this._annotations;
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to set the main body part of a Word file
   * @param {string} body the body string 
   */
  setBody(body) {
    this._body = body;
  }

  /**
   * Accessor to set the footnotes part of a Word file
   * @param {string} footnotes the footnotes string 
   */
  setFootnotes(footnotes) {
    this._footnotes = footnotes;
  }

  /**
   * Accessor to set the endnotes part of a Word file
   * @param {string} endnotes the endnotes string 
   */
  setEndnotes(endnotes) {
    this._endnotes = endnotes;
  }

  /**
   * Accessor to set the headers part of a Word file
   * @param {string} headers the headers string 
   */
  setHeaders(headers) {
    this._headers = headers;
  }

  /**
   * Accessor to set the footers part of a Word file
   * @param {string} footers the footers string 
   */
  setFooters(footers) {
    this._footers = footers;
  }

  /**
   * Accessor to set the annotations part of a Word file
   * @param {string} annotations the annotations string 
   */
  setAnnotations(annotations) {
    this._annotations = annotations;
  }
}


module.exports = Document;

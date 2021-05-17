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

class Document {
  
  constructor() {
    this._body = "";
    this._footnotes = "";
    this._endnotes = "";
    this._headers = "";
    this._annotations = "";
  }

  /**
   * Accessor to read the main body part of a Word file
   * @returns a string, containing the Word file body
   */
  getBody(filterUnicode) { 
    const value = this._body;
    return (filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the footnotes part of a Word file
   * @returns a string, containing the Word file footnotes
   */
  getFootnotes(filterUnicode) {
    const value = this._footnotes;
    return (filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the endnotes part of a Word file
   * @returns a string, containing the Word file endnotes
   */
  getEndnotes(filterUnicode) {
    const value = this._endnotes;
    return (filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the headers part of a Word file
   * @returns a string, containing the Word file headers
   */
  getHeaders(filterUnicode) {
    const value = this._headers;
    return (filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the annotations part of a Word file
   * @returns a string, containing the Word file annotations
   */
  getAnnotations(filterUnicode) {
    const value = this._annotations;
    return (filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to set the main body part of a Word file
   * @param {*} body the body string 
   */
  setBody(body) {
    this._body = body;
  }

  /**
   * Accessor to set the footnotes part of a Word file
   * @param {*} footnotes the footnotes string 
   */
  setFootnotes(footnotes) {
    this._footnotes = footnotes;
  }

  /**
   * Accessor to set the endnotes part of a Word file
   * @param {*} endnotes the endnotes string 
   */
  setEndnotes(endnotes) {
    this._endnotes = endnotes;
  }

  /**
   * Accessor to set the headers part of a Word file
   * @param {*} headers the headers string 
   */
  setHeaders(headers) {
    this._headers = headers;
  }

  /**
   * Accessor to set the annotations part of a Word file
   * @param {*} annotations the annotations string 
   */
  setAnnotations(annotations) {
    this._annotations = annotations;
  }
}


module.exports = Document;

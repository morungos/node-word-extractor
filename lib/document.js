/**
 * @module lib/document
 * 
 * @description
 * Implements the main document returned when a Word file has been extracted. This exposes
 * methods that allow the body, annotations, headers, footnotes, and endnotes, to be 
 * read and used.
 * 
 * @author
 * Stuart Watt <stuart@morungos.com>
 */

class Document {
  
  constructor() {}

  /**
   * Accessor to read the main body part of a Word file
   * @returns a string, containing the Word file body
   */
  getBody() { 
    return this._body;
  }

  /**
   * Accessor to read the footnotes part of a Word file
   * @returns a string, containing the Word file footnotes
   */
  getFootnotes() {
    return this._footnotes;
  }

  /**
   * Accessor to read the endnotes part of a Word file
   * @returns a string, containing the Word file endnotes
   */
  getEndnotes() {
    return this._endnotes;
  }

  /**
   * Accessor to read the headers part of a Word file
   * @returns a string, containing the Word file headers
   */
  getHeaders() {
    return this._headers;
  }

  /**
   * Accessor to read the annotations part of a Word file
   * @returns a string, containing the Word file annotations
   */
  getAnnotations() {
    return this._annotations;
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

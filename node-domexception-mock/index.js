if (!globalThis.DOMException) {
  // Polyfill for extremely old environments if any (fallback)
  globalThis.DOMException = class DOMException extends Error {
    constructor(message, name) {
      super(message);
      this.name = name || "DOMException";
    }
  };
}
module.exports = globalThis.DOMException;

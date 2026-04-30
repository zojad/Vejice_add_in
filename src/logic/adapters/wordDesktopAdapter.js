export class WordDesktopAdapter {
  constructor({ textBridge, trace = () => {} }) {
    this.textBridge = textBridge;
    this.trace = trace;
  }

  async getParagraphCollection(context) {
    // Win32 Word can fail on direct body.paragraphs in some document states.
    // Using body range first is more stable across builds.
    const paras = context.document.body.getRange().paragraphs;
    this.trace("DesktopAdapter:getParagraphs load(items) -> sync:start");
    paras.load("items");
    await context.sync();
    this.trace("DesktopAdapter:getParagraphs load(items) -> sync:done", paras.items.length);
    return paras;
  }

  async loadParagraphTexts(context, paras, indexes = null) {
    if (!paras?.items?.length) return;
    let targets = paras.items;
    if (Array.isArray(indexes) && indexes.length) {
      const uniqueIndexes = [...new Set(indexes)].filter(
        (value) => Number.isFinite(value) && value >= 0 && value < paras.items.length
      );
      targets = uniqueIndexes.map((idx) => paras.items[idx]).filter(Boolean);
    }
    if (!targets.length) return;
    this.trace("DesktopAdapter:getParagraphs load(item.text) -> sync:start");
    targets.forEach((p) => p.load("text"));
    await context.sync();
    this.trace("DesktopAdapter:getParagraphs load(item.text) -> sync:done", targets.length);
  }

  async getParagraphs(context) {
    // Desktop Word can throw on shorthand nested loads (e.g. "items/text"),
    // so load items first, then load each paragraph text explicitly.
    const paras = await this.getParagraphCollection(context);
    await this.loadParagraphTexts(context, paras);
    return paras;
  }

  async applySuggestion(context, paragraph, suggestion) {
    return this.textBridge.applySuggestion(context, paragraph, suggestion);
  }
}

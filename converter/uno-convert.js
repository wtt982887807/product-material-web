/** Convert an in-memory file using real LibreOffice UNO, never a text rebuild. */
export function convertUnoDocument({inputUrl, outputUrl}, {css, zetajs, desktop, context}) {
  const property = (Name, Value) => new css.beans.PropertyValue({Name, Value});
  const short = value => new zetajs.Any(zetajs.type.short, value);
  const interactionHandler = zetajs.unoObject([css.task.XInteractionHandler], {
    handle(request) {
      for (const continuation of request.getContinuations()) {
        const abort = continuation.queryInterface(zetajs.type.interface(css.task.XInteractionAbort));
        if (abort) { abort.select(); return; }
      }
      throw new Error('文件要求交互确认或密码，无法自动转换。');
    },
  });
  let model;
  try {
    model = desktop.loadComponentFromURL(inputUrl, '_blank', 0, [
      property('Hidden', true), property('ReadOnly', false), property('Silent', true),
      property('MacroExecutionMode', short(0)), property('UpdateDocMode', short(0)),
      property('InteractionHandler', interactionHandler),
    ]);
    if (!model?.supportsService('com.sun.star.text.TextDocument')) throw new Error('文件未被识别为可转换的 Word 文档。');
    if (typeof model.getRedlines !== 'function') throw new Error('无法检查已有 Word 修订，已停止转换。');
    const redlines = model.getRedlines();
    if (redlines.getCount() > 0) {
      const frame = model.getCurrentController()?.getFrame();
      if (!frame) throw new Error('无法接受已有 Word 修订，已停止转换。');
      css.frame.DispatchHelper.create(context).executeDispatch(frame, '.uno:AcceptAllTrackedChanges', '', 0, []);
    }
    if (model.getRedlines().getCount() !== 0) throw new Error('仍有未能接受的旧修订，已停止转换，请先在 Word 中处理。');
    model.storeToURL(outputUrl, [property('Overwrite', true), property('FilterName', 'Office Open XML Text')]);
  } finally {
    if (model) { try { model.close(true); } catch { try { model.dispose(); } catch {} } }
  }
}

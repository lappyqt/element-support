/**
 * @OnlyCurrentDoc
 */
function onOpen() {
  SupportMenu.addMenuToTable();
}

/**
 * @OnlyCurrentDoc
 */
function onEdit(e) {
  SupportMenu.setReprimandNotes(e);
}
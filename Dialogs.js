/**
 * Открывает выбор шаблона для Prepare Template.
 */
function showPrepareTemplateDialog() {
  showTemplateSelector_('prepare');
}


/**
 * Открывает выбор шаблона для Update Payees.
 */
function showUpdatePayeesDialog() {
  showTemplateSelector_('payees');
}


/**
 * Показывает dropdown с доступными шаблонами.
 */
function showTemplateSelector_(action) {
  const template = HtmlService.createTemplateFromFile(
    'TemplateSelector'
  );

  template.action = action;
  template.templateKeys = CONFIG.TEMPLATE_KEYS;

  const html = template
    .evaluate()
    .setWidth(320)
    .setHeight(190);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    action === 'prepare'
      ? 'Prepare Template'
      : 'Update Payees'
  );
}


/**
 * Получает выбор из HTML-окна
 * и запускает нужное действие.
 */
function runTemplateAction(action, templateKey) {
  validateTemplateKey_(templateKey);

  switch (action) {
    case 'prepare':
      prepareTemplate(templateKey);
      break;

    case 'payees':
      updatePayees(templateKey);
      break;

    default:
      throw new Error(
        `Unknown action: ${action}`
      );
  }
}
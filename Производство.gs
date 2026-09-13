function handleProduction(ss, sheet, range, row) {
  const context = resolveActionContext_(ss, sheet, row);
  if (!context) {
    Logger.log(`handleProduction: не удалось определить строку База для ${sheet.getName()}!${row}.`);
    return null;
  }

  const current = readBaseFields_(context.baseSheet, context.baseRow, ['stage_production']).stage_production;
  if (current) {
    resetRoleCheckbox_(context, range);
    return context.itemId;
  }

  resetRoleCheckbox_(context, range);
  // Завершение производства должно переводить изделие НА этап ОТК, а не
  // автоматически завершать и ОТК. stage_quality_check отмечается только на ОТК.
  writeBaseFields_(context.baseSheet, context.baseRow, { stage_production: new Date() });
  return context.itemId;
}

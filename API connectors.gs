const SALESDRIVE_FORM_ID = '2HqZ95-RMpEmW2_pOzxtQAkuYrrohkZVLMmtznjzfro0rUNBpiOty5xiQzy6';

/** Экспортирует состав заказа в SalesDrive и возвращает ID созданного заказа. */
function sendPostRequest(products, comment) {
  const payload = {
    form: SALESDRIVE_FORM_ID,
    getResultData: '1',
    products,
    comment,
    externalId: '',
    fName: '',
    lName: '',
    mName: '',
    phone: '',
    email: '',
    con_comment: '',
    shipping_method: '',
    payment_method: '',
    shipping_address: '',
    novaposhta: {
      ServiceType: '', payer: '', area: '', city: '', cityNameFormat: '',
      WarehouseNumber: '', Street: '', BuildingNumber: '', Flat: ''
    },
    ukrposhta: {
      ServiceType: '', payer: '', city: '', WarehouseNumber: '',
      Street: '', BuildingNumber: '', Flat: ''
    },
    justin: { WarehouseNumber: '' },
    sajt: '',
    con_telegram: '',
    organizationId: '',
    shipping_costs: '',
    tTN: '',
    con_conrivenCin2: '',
    kanalTrafiku: '',
    tipKontragenta: '',
    proekt: '',
    con_consposibDostavki: '',
    con_condostavka: '',
    zvorotnijZvazok: '',
    ktoPakovalPosylku: '',
    stockId: ''
  };

  const response = UrlFetchApp.fetch('https://acs.salesdrive.me/handler/', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });

  const responseObject = JSON.parse(response.getContentText());
  return responseObject && responseObject.data ? responseObject.data.orderId : null;
}

/** Меняет статус существующего заказа SalesDrive. */
function changeSdStatus(id, statusId) {
  if (!id) {
    Logger.log('changeSdStatus: отсутствует ID заказа SalesDrive, запрос пропущен.');
    return false;
  }

  const payload = {
    form: SALESDRIVE_FORM_ID,
    id,
    externalId: '',
    data: {
      statusId,
      paymentDate: '',
      rejectionReasonId: '',
      salesdrive_manager: '',
      comment: '',
      fName: '',
      lName: '',
      mName: '',
      phone: '',
      email: '',
      con_comment: '',
      shipping_method: '',
      payment_method: '',
      shipping_address: '',
      novaposhta: { ttn: '' },
      ukrposhta: { ttn: '' },
      sajt: '',
      con_telegram: '',
      organizationId: '',
      shipping_costs: '',
      tTN: '',
      con_conrivenCin2: '',
      kanalTrafiku: '',
      tipKontragenta: '',
      proekt: '',
      con_consposibDostavki: '',
      con_condostavka: '',
      zvorotnijZvazok: '',
      ktoPakovalPosylku: ''
    }
  };

  const response = UrlFetchApp.fetch('https://acs.salesdrive.me/api/order/update/', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const ok = response.getResponseCode() >= 200 && response.getResponseCode() < 300;
  if (!ok) Logger.log('SalesDrive status update error: ' + response.getContentText());
  return ok;
}

/**
 * Единая точка хранения параметров Trello внутри кода.
 * Отдельным следующим шагом их стоит перенести в Script Properties и перевыпустить
 * token, поскольку он уже присутствует в истории Git.
 */
function getTrelloCredentials_() {
  return {
    apiKey: 'c15d1e1174e2b9d45b2fc45629c563e9',
    token: 'ATTA8289839f0b080402baa34f55c6a8b1118d64b693ab64ec78c91a579b528954786E0863A2',
    boardId: 'a70fjjsT',
    listId: '644394d259dde2c026d4c8aa'
  };
}

/** Создаёт карточку Trello для производственной задачи. */
function createTrelloCard(name, qty, destination) {
  const credentials = getTrelloCredentials_();
  const cleanDestination = normalizeDestination(destination);
  const cardName = `${name} = ${qty} (${cleanDestination})`;

  const url = 'https://api.trello.com/1/cards?key=' + encodeURIComponent(credentials.apiKey) +
    '&token=' + encodeURIComponent(credentials.token) +
    '&idList=' + encodeURIComponent(credentials.listId) +
    '&name=' + encodeURIComponent(cardName) +
    '&desc=&idBoard=' + encodeURIComponent(credentials.boardId) +
    '&pos=top';

  const response = UrlFetchApp.fetch(url, { method: 'post' });
  const responseData = JSON.parse(response.getContentText());
  return responseData.id;
}

/** Архивирует карточку Trello. */
function closeTrelloCard(cardId) {
  if (!cardId) return false;

  const credentials = getTrelloCredentials_();
  const archiveUrl = `https://api.trello.com/1/cards/${encodeURIComponent(cardId)}?closed=true&key=${encodeURIComponent(credentials.apiKey)}&token=${encodeURIComponent(credentials.token)}`;

  try {
    const response = UrlFetchApp.fetch(archiveUrl, {
      method: 'put',
      muteHttpExceptions: true
    });

    if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
      return true;
    }

    Logger.log('Не удалось архивировать карточку Trello: ' + response.getContentText());
    return false;
  } catch (error) {
    Logger.log('Ошибка Trello: ' + error.message);
    return false;
  }
}

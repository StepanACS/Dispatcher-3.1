//Экспорт спецификации в SalesDrive
function sendPostRequest(products, comment) {
  var url = 'https://acs.salesdrive.me/handler/';
  var payload = {
    "form": "2HqZ95-RMpEmW2_pOzxtQAkuYrrohkZVLMmtznjzfro0rUNBpiOty5xiQzy6",
    "getResultData": "1",
    "products": products,
    "comment": comment,
    "externalId": "",
    "fName": "",
    "lName": "",
    "mName": "",
    "phone": "",
    "email": "",
    "con_comment": "",
    "shipping_method": "",
    "payment_method": "",
    "shipping_address": "",
    "novaposhta": {
        "ServiceType": "",
        "payer": "",
        "area": "",
        "city": "",
        "cityNameFormat": "",
        "WarehouseNumber": "",
        "Street": "",
        "BuildingNumber": "",
        "Flat": ""
    },
    "ukrposhta": {
        "ServiceType": "",
        "payer": "",
        "city": "",
        "WarehouseNumber": "",
        "Street": "",
        "BuildingNumber": "",
        "Flat": ""
    },
    "justin": {
        "WarehouseNumber": ""
    },
    "sajt": "",
    "con_telegram": "",
    "organizationId": "",
    "shipping_costs": "",
    "tTN": "",
    "con_conrivenCin2": "",
    "kanalTrafiku": "",
    "tipKontragenta": "",
    "proekt": "",
    "con_consposibDostavki": "",
    "con_condostavka": "",
    "zvorotnijZvazok": "",
    "ktoPakovalPosylku": "",
    "stockId": ""
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
  var responseObject = JSON.parse(response); // Преобразуем строку JSON в объект JavaScript
  var orderId = responseObject.data.orderId; // Получаем значение orderId из объекта
  return orderId;
}

function changeSdStatus(id, statusId) {  //запрос на изменение статуса заказа
  var url = 'https://acs.salesdrive.me/api/order/update/';
  var payload = {
  "form": "2HqZ95-RMpEmW2_pOzxtQAkuYrrohkZVLMmtznjzfro0rUNBpiOty5xiQzy6",
  "id": id,
  "externalId": "",
  "data": {
    "statusId": statusId, //23 - "Передано в работу"
    "paymentDate": "",
    "rejectionReasonId": "",
    "salesdrive_manager": "",
    "comment": "",
    "fName": "",
    "lName": "",
    "mName": "",
    "phone": "",
    "email": "",
    "con_comment": "",
    "shipping_method": "",
    "payment_method": "",
    "shipping_address": "",
    "novaposhta": {
      "ttn": ""
    },
    "ukrposhta": {
      "ttn": ""
    },
    "sajt": "",
    "con_telegram": "",
    "organizationId": "",
    "shipping_costs": "",
    "tTN": "",
    "con_conrivenCin2": "",
    "kanalTrafiku": "",
    "tipKontragenta": "",
    "proekt": "",
    "con_consposibDostavki": "",
    "con_condostavka": "",
    "zvorotnijZvazok": "",
    "ktoPakovalPosylku": ""
    }
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}


function createProduct() {  //запрос на изменение статуса заказа
  var url = 'https://acs.salesdrive.me/product-handler/';
  var payload = {

    "form": "2HqZ95-RMpEmW2_pOzxtQAkuYrrohkZVLMmtznjzfro0rUNBpiOty5xiQzy6",
    "action": "update",
    //"dontUpdateFields" => [],
    "product": [
        {
            "id": "",
            "name": "",
            "nameTranslate": "",
            "descriptionTranslate": "",
            "nameForDocuments": "",
            "costPerItem": "",
            "sku": "",
            "manufacturer": "",
            "currency": "",
            "discount": {
                "value": "",
                "date_start": "",
                "date_end": ""
            },
            "weight": "",
            "volume": "",
            "length": "",
            "width": "",
            "height": "",
            "barcode": "",
            "uktzed": "",
            "stockBalanceByStock": {
                "3": "", 
                "5": ""
            },
            "expenses": "",
            "currencyExpenses": "",
            "category": {
                "id": "",
                "name": ""
            },
            "description": "",
            "url": "",
            "note": "",
            "supplier": "",
            "keywords": "",
            "images": [
                {
                    "fullsize": "",
                    "thumbnail": ""
                },
                {
                    "fullsize": "",
                    "thumbnail": ""
                }
            ],
            "params": [
                {
                    "name": "",
                    "type": "",
                    "value": ""
                },
                {
                    "name": "",
                    "type": "",
                    "value": ""
                }
            ],
            "parentProductId": "",
            "additionalPrices": [
                {
                    "priceType": "",
                    "priceValue": "",
                    "priceCurrency": "",
                    "priceDiscount": ""
                },
                {
                    "priceType": "",
                    "priceValue": "",
                    "priceCurrency": "",
                    "priceDiscount": ""
                }
            ],
            "label": []
        }
    ]
}
}

//Создание карточки в трелло


function createTrelloCard(name, qty,destination ) {
  destination = normalizeDestination(destination);
  var apiKey = 'c15d1e1174e2b9d45b2fc45629c563e9';
  //var token = 'ATTA3be61fedb1bf70893d7d18021165f3a4d9e8e511257d07fb4d9b7acaccc9cd9295401734';
  const token = 'ATTA8289839f0b080402baa34f55c6a8b1118d64b693ab64ec78c91a579b528954786E0863A2';
  
  var boardId = 'a70fjjsT';
  var listId = '644394d259dde2c026d4c8aa';
  var cardName = name + " = " + qty + " (" + destination + ")"; 
  var cardDescription = '';

  var url = 'https://api.trello.com/1/cards?key=' + apiKey + '&token=' + token +
            '&idList=' + listId + '&name=' + encodeURIComponent(cardName) +
            '&desc=' + encodeURIComponent(cardDescription) + '&idBoard=' + boardId +
            '&pos=top'; 

  var options = {
    'method': 'post'
  };

  var response = UrlFetchApp.fetch(url, options);
  var responseData = JSON.parse(response.getContentText()); // Преобразовать текст ответа в объект JSON
  var cardId = responseData.id; // Получить ID карточки из объекта ответа
  Logger.log('ID созданной карточки:', cardId);
  return cardId;
}

function closeTrelloCard(cardId) {
  var ui = SpreadsheetApp.getUi();
  try {
    // Ваши учетные данные для доступа к Trello API
    const apiKey = 'c15d1e1174e2b9d45b2fc45629c563e9';
    const token = 'ATTA8289839f0b080402baa34f55c6a8b1118d64b693ab64ec78c91a579b528954786E0863A2';
  
    // URL для запроса на архивацию карточки
    const archiveUrl = `https://api.trello.com/1/cards/${cardId}?closed=true&key=${apiKey}&token=${token}`;

    // Отправляем запрос на архивацию карточки
    var response = UrlFetchApp.fetch(archiveUrl, {
      method: 'PUT',
    });

    if (response.getResponseCode() === 200) {
      console.log('Card archived successfully');
      //var message = ui.alert("Карточка в Трелло найдена и успешно архивирована  ", ui.ButtonSet.OK);
    } else {
      console.error('Failed to archive card');
      var message = ui.alert("Не удалось архивировать карточку в Трелло, выполните эту работу вручную  ", ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('An error occurred:', error);
    var message = ui.alert("Произошла ошибка при обращении к Trello API. Пожалуйста, выполните эту работу вручную. ", ui.ButtonSet.OK);
  }
}

function trelloAuthCheck() {
  const apiKey = 'c15d1e1174e2b9d45b2fc45629c563e9';
  //var token  = 'ATTA3be61fedb1bf70893d7d18021165f3a4d9e8e511257d07fb4d9b7acaccc9cd9295401734';
  const token = 'ATTA8289839f0b080402baa34f55c6a8b1118d64b693ab64ec78c91a579b528954786E0863A2';
  

  var url = 'https://api.trello.com/1/members/me?key=' + encodeURIComponent(apiKey) +
            '&token=' + encodeURIComponent(token);

  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log(resp.getResponseCode());
  Logger.log(resp.getContentText());
}




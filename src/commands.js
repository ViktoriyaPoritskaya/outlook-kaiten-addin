/* global Office, window, console */

// ВАЖНО: ES5-совместимый код (var, function, Promise-цепочки вместо async/await)
// — Outlook 2016 desktop выполняет надстройку в движке IE11.

Office.onReady(function () {
  if (typeof Office.actions !== "undefined" && Office.actions.associate) {
    Office.actions.associate("createKaitenCard", createKaitenCard);
  }
});

/**
 * Точка входа кнопки «Создать задачу» в группе Kaiten.
 * Перед созданием показывает окно выбора доски: доска по умолчанию + возможность
 * выбрать другую только для этого письма (или отменить).
 */
function createKaitenCard(event) {
  var item = Office.context.mailbox.item;

  var token = window.KaitenSettings.getToken();
  if (!token) {
    notify(item, "kaiten-error", "Не настроен токен. Открой «Настройки» в группе Kaiten.", "errorMessage");
    event.completed();
    return;
  }
  var defaultBoardId = window.KaitenSettings.getBoardId();
  if (!defaultBoardId) {
    notify(item, "kaiten-error", "Не выбрана доска по умолчанию. Открой «Настройки» в группе Kaiten.", "errorMessage");
    event.completed();
    return;
  }
  var defaultBoardTitle = window.KaitenSettings.getBoardTitle();

  chooseBoard(defaultBoardId, defaultBoardTitle)
    .then(function (choice) {
      if (!choice) {
        // Пользователь нажал «Отмена» или закрыл окно.
        notify(item, "kaiten-info", "Создание задачи отменено.", "informationalMessage");
        event.completed();
        return;
      }

      notify(item, "kaiten-info", "Создаю задачу в Kaiten…", "informationalMessage");
      return run(item, choice.boardId).then(function (card) {
        var url = window.KaitenApi.getCardUrl(card.id);
        notify(item, "kaiten-info", "Задача создана: #" + card.id, "informationalMessage");

        // Статус вложений — отдельным уведомлением, чтобы не обрезался лимитом длины.
        var att = card.__attachments;
        if (att) {
          if (att.unsupported) {
            notify(item, "kaiten-att", "Вложения: этот Outlook не отдаёт файлы (нет ни API 1.8, ни EWS).", "informationalMessage");
          } else if (att.total) {
            var attMsg = "Вложения (" + att.method + "): загружено " + att.uploaded + " из " + att.total + ".";
            if (att.uploaded !== att.total && att.lastError) {
              attMsg += " Причина: " + att.lastError;
            }
            notify(
              item,
              "kaiten-att",
              attMsg,
              att.uploaded === att.total ? "informationalMessage" : "errorMessage"
            );
          }
        }
        if (window.KAITEN_CONFIG.OPEN_CARD_AFTER_CREATE) {
          openInBrowser(url);
        }
        event.completed();
      });
    })
    .catch(function (err) {
      if (console && console.error) console.error("[Kaiten Add-in]", err);
      notify(item, "kaiten-error", short(err && err.message ? err.message : String(err)), "errorMessage");
      event.completed();
    });
}

/**
 * Готовит список досок и показывает окно выбора.
 * Возвращает Promise с { boardId, boardTitle } либо null (отмена).
 */
function chooseBoard(defaultBoardId, defaultBoardTitle) {
  return window.KaitenApi
    .listBoards()
    .catch(function () {
      // Не удалось загрузить список — покажем окно только с доской по умолчанию.
      return [];
    })
    .then(function (boards) {
      var list = normalizeBoards(boards);
      var payload = {
        boards: list,
        defaultBoardId: defaultBoardId,
        defaultBoardTitle: defaultBoardTitle || "",
      };
      try {
        window.localStorage.setItem("kaiten.boardChoice", JSON.stringify(payload));
      } catch (e) {
        /* нет localStorage — окно возьмёт дефолт из query-параметров */
      }
      return openBoardDialog(defaultBoardId, defaultBoardTitle);
    });
}

// Приводит ответ listBoards к массиву { id, title }.
function normalizeBoards(boards) {
  var arr = Array.isArray(boards) ? boards : (boards && boards.data) || [];
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    out.push({ id: arr[i].id, title: arr[i].title || ("Доска #" + arr[i].id) });
  }
  return out;
}

// Открывает dialog.html через Office Dialog API. Резолвит выбор пользователя.
function openBoardDialog(defaultBoardId, defaultBoardTitle) {
  return new Promise(function (resolve, reject) {
    if (!Office.context.ui || !Office.context.ui.displayDialogAsync) {
      // Совсем старый клиент без Dialog API — создаём сразу на доске по умолчанию.
      resolve({ boardId: defaultBoardId, boardTitle: defaultBoardTitle || "" });
      return;
    }

    // _ts — защита от кэша: Outlook/IE кэширует окно диалога отдельно от панели,
    // поэтому каждый раз запрашиваем свежую страницу с уникальным адресом.
    var url =
      dialogUrl() +
      "?defBoard=" +
      encodeURIComponent(defaultBoardId) +
      "&defTitle=" +
      encodeURIComponent(defaultBoardTitle || "") +
      "&_ts=" +
      new Date().getTime();

    Office.context.ui.displayDialogAsync(
      url,
      { height: 48, width: 34, displayInIframe: false },
      function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Не удалось открыть окно выбора доски: " + (res.error && res.error.message)));
          return;
        }
        var dialog = res.value;
        var settled = false;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          var msg = null;
          try { msg = JSON.parse(arg.message); } catch (e) { /* ignore */ }
          settled = true;
          dialog.close();
          if (msg && msg.action === "confirm") {
            resolve({ boardId: Number(msg.boardId) || 0, boardTitle: msg.boardTitle || "" });
          } else {
            resolve(null);
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
          // Пользователь закрыл окно крестиком и т.п. — считаем отменой.
          if (!settled) resolve(null);
        });
      }
    );
  });
}

// URL окна dialog.html рядом с commands.html (тот же каталог).
function dialogUrl() {
  return window.location.href.replace(/[^/]*$/, "dialog.html");
}

/**
 * Основной сценарий создания карточки на заданной доске. Возвращает Promise с карточкой.
 */
function run(item, boardId) {
  if (!boardId) {
    return Promise.reject(new Error("Не выбрана доска для задачи."));
  }

  var card;
  var data;
  var customerPropId = null;
  var customerValue = "";

  return collectEmailData(item)
    .then(function (collected) {
      data = collected;
      customerValue = formatSender(data);
      // Список кастомных полей нужен, чтобы найти id поля «Заказчик».
      // Best effort: если не получится — просто не заполним это поле.
      return window.KaitenApi.listCustomProperties().catch(function () {
        return [];
      });
    })
    .then(function (props) {
      customerPropId = findPropertyId(normalizeProps(props), customerFieldName());

      var description = buildDescription(data, window.KAITEN_CONFIG);

      var payload = {
        board_id: boardId,
        title: truncateLine(data.subject, 250),
        description: description,
      };

      var columnId = window.KaitenSettings.getDefaultColumnId();
      if (columnId) payload.column_id = columnId;

      return window.KaitenApi.createCard(payload);
    })
    .then(function (created) {
      card = created;

      // Дописываем поле «Заказчик» отдельным запросом. Best effort: если не выйдет —
      // карточка уже создана, ошибку только залогируем, но не роняем весь сценарий.
      if (customerPropId && customerValue) {
        var props = {};
        props["id_" + customerPropId] = customerValue;
        return window.KaitenApi
          .updateCard(card.id, { properties: props })
          .catch(function (e) {
            if (console && console.warn) {
              console.warn("[Kaiten Add-in] Не удалось заполнить поле «Заказчик»:", e);
            }
          });
      }
      return null;
    })
    .then(function () {
      // Загружаем сами файлы вложений в карточку (best effort). Сводку кладём на card.
      return uploadAttachments(item, card.id)
        .then(function (summary) {
          card.__attachments = summary;
          // Диагностика вложений — комментарием в карточку (надёжнее исчезающих уведомлений).
          if (!summary || summary.uploaded !== summary.total || summary.unsupported) {
            return window.KaitenApi
              .addComment(card.id, buildAttachmentDiag(item, summary))
              .catch(function () { /* диагностика — best effort */ });
          }
          return null;
        })
        .catch(function (e) {
          if (console && console.warn) {
            console.warn("[Kaiten Add-in] Не удалось загрузить вложения:", e);
          }
        });
    })
    .then(function () {
      // Прикрепляем ссылку на исходное письмо как external link (best effort).
      try {
        if (item.itemId && Office.context.mailbox.convertToRestId) {
          var restId = Office.context.mailbox.convertToRestId(
            item.itemId,
            Office.MailboxEnums.RestVersion.v2_0
          );
          var link =
            "https://outlook.office.com/owa/?ItemID=" +
            encodeURIComponent(restId) +
            "&exvsurl=1&viewmodel=ReadMessageItem";
          return window.KaitenApi
            .addExternalLink(card.id, link, "Исходное письмо в Outlook")
            .catch(function (e) {
              if (console && console.warn) {
                console.warn("[Kaiten Add-in] Не удалось добавить ссылку на письмо:", e);
              }
            });
        }
      } catch (e) {
        if (console && console.warn) {
          console.warn("[Kaiten Add-in] Не удалось добавить ссылку на письмо:", e);
        }
      }
      return null;
    })
    .then(function () {
      return card;
    });
}

/**
 * Собирает данные из открытого письма (тело — асинхронно).
 */
function collectEmailData(item) {
  return new Promise(function (resolve, reject) {
    var data = {
      subject: item.subject || "(без темы)",
      senderName: "",
      senderEmail: "",
      body: "",
      attachments: [],
      receivedDate: item.dateTimeCreated ? new Date(item.dateTimeCreated) : null,
    };

    var sender = item.from || item.sender;
    if (sender) {
      data.senderName = sender.displayName || "";
      data.senderEmail = sender.emailAddress || "";
    }

    if (Array.isArray(item.attachments)) {
      data.attachments = item.attachments.map(function (a) {
        return {
          name: a.name || "(без имени)",
          size: typeof a.size === "number" ? a.size : null,
        };
      });
    }

    item.body.getAsync(Office.CoercionType.Text, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        data.body = result.value || "";
        resolve(data);
      } else {
        reject(new Error("Не удалось прочитать тело письма: " + (result.error && result.error.message)));
      }
    });
  });
}

/**
 * Markdown-описание для карточки Kaiten (Kaiten поддерживает Markdown в описании).
 */
function buildDescription(data, cfg) {
  var lines = [];

  if (cfg.INCLUDE_SENDER_IN_DESCRIPTION && (data.senderName || data.senderEmail)) {
    var sender = data.senderName
      ? data.senderName + (data.senderEmail ? " <" + data.senderEmail + ">" : "")
      : data.senderEmail;
    lines.push("**От:** " + sender);
  }

  if (cfg.INCLUDE_RECEIVED_DATE_IN_DESCRIPTION && data.receivedDate) {
    lines.push("**Дата:** " + formatDate(data.receivedDate));
  }

  if (cfg.INCLUDE_ATTACHMENTS_IN_DESCRIPTION && data.attachments.length > 0) {
    lines.push("**Вложения:**");
    data.attachments.forEach(function (a) {
      lines.push(
        "- " + a.name + (a.size != null ? " _(" + formatBytes(a.size) + ")_" : "")
      );
    });
  }

  if (lines.length > 0) {
    lines.push("");
    lines.push("---");
    lines.push("");
  }

  lines.push(truncate(data.body, cfg.MAX_BODY_LENGTH));

  return lines.join("\n");
}

// Проверяет, умеет ли текущий Outlook отдавать содержимое вложений
// (getAttachmentContentAsync появился в Mailbox 1.8).
function supportsAttachmentContent(item) {
  try {
    if (Office.context.requirements && Office.context.requirements.isSetSupported) {
      if (!Office.context.requirements.isSetSupported("Mailbox", "1.8")) return false;
    }
    return typeof item.getAttachmentContentAsync === "function";
  } catch (e) {
    return false;
  }
}

// Загружает все файловые вложения письма в карточку. Best effort, последовательно.
// Способы пробуются по очереди, пока какой-то не отдаст содержимое:
//   1) getAttachmentContentAsync (Mailbox 1.8) — современные клиенты;
//   2) Outlook REST API (getCallbackTokenAsync, с 1.5);
//   3) EWS GetItem + MimeContent (с 1.5) — разбор письма в MIME на месте.
// Возвращает Promise со сводкой { uploaded, total, method, lastError, unsupported }.
function uploadAttachments(item, cardId) {
  var atts = [];
  var all = item.attachments || [];
  for (var i = 0; i < all.length; i++) {
    var a = all[i];
    // Берём только файлы (не вложенные письма/облачные ссылки).
    if (!a.attachmentType || a.attachmentType === "file") atts.push(a);
  }
  if (!atts.length) return Promise.resolve({ uploaded: 0, total: 0 });

  var methods = [];
  if (supportsAttachmentContent(item)) {
    methods.push({ name: "1.8", fn: function () { return collectVia18(item, atts); } });
  }
  if (canUseRest(item)) {
    methods.push({ name: "REST", fn: function () { return collectViaRest(item); } });
  }
  if (canUseEws(item)) {
    methods.push({ name: "MIME", fn: function () { return collectViaMime(item, atts); } });
  }
  if (!methods.length) {
    return Promise.resolve({ uploaded: 0, total: atts.length, unsupported: true });
  }

  return tryCollectors(methods, atts, cardId, 0, "");
}

// Перебор способов получения содержимого: при пустом результате/ошибке пробуем следующий.
function tryCollectors(methods, atts, cardId, idx, errLog) {
  if (idx >= methods.length) {
    return Promise.resolve({
      uploaded: 0,
      total: atts.length,
      method: "—",
      lastError: errLog || "не удалось получить содержимое",
    });
  }
  var mth = methods[idx];
  return mth.fn().then(function (res) {
    if (res.error || !res.files || !res.files.length) {
      var e = res.error || "нет файлов с содержимым";
      var nextLog = errLog + (errLog ? "; " : "") + mth.name + ": " + e;
      return tryCollectors(methods, atts, cardId, idx + 1, nextLog);
    }
    return uploadFileList(res.files, cardId).then(function (u) {
      var note = errLog ? " (до этого — " + errLog + ")" : "";
      return {
        uploaded: u.uploaded,
        total: res.files.length,
        method: mth.name,
        lastError: (u.lastError || "") + note,
      };
    });
  });
}

// Последовательная загрузка готового списка файлов { name, contentType, base64 }.
function uploadFileList(files, cardId) {
  var uploaded = 0;
  var lastError = "";
  var chain = Promise.resolve();
  files.forEach(function (f) {
    chain = chain.then(function () {
      if (!f.base64) {
        lastError = "нет содержимого: " + (f.name || "?");
        return null;
      }
      var blob = base64ToBlob(f.base64, f.contentType);
      return window.KaitenApi
        .uploadCardFile(cardId, blob, f.name || "attachment")
        .then(function () { uploaded++; })
        .catch(function (e) {
          lastError = (e && e.message) ? e.message : String(e);
          if (console && console.warn) {
            console.warn("[Kaiten Add-in] Вложение не загружено:", f.name, e);
          }
        });
    });
  });
  return chain.then(function () {
    return { uploaded: uploaded, lastError: lastError };
  });
}

// Диагностический текст для комментария в карточку (пока отлаживаем вложения).
function buildAttachmentDiag(item, summary) {
  var lines = ["Диагностика вложений (служебное):"];
  var all = item.attachments || [];
  lines.push("- всего в письме: " + all.length);
  for (var i = 0; i < all.length; i++) {
    var a = all[i];
    lines.push(
      "  • " + (a.name || "?") +
      " [type=" + (a.attachmentType || "?") +
      ", size=" + (a.size != null ? a.size : "?") +
      ", inline=" + (a.isInline ? "да" : "нет") + "]"
    );
  }
  if (summary) {
    if (summary.unsupported) {
      lines.push("- способ: недоступен (ни Mailbox 1.8, ни REST)");
    } else {
      lines.push("- способ: " + (summary.method || "?"));
      lines.push("- загружено: " + summary.uploaded + " из " + summary.total);
      if (summary.lastError) lines.push("- причина: " + summary.lastError);
    }
  }
  return lines.join("\n");
}

// Способ 1: собрать файлы через getAttachmentContentAsync (Mailbox 1.8).
// Возвращает { files:[{name,contentType,base64}], error }.
function collectVia18(item, atts) {
  var files = [];
  var lastErr = "";
  var chain = Promise.resolve();
  atts.forEach(function (att) {
    chain = chain.then(function () {
      return new Promise(function (resolve) {
        try {
          item.getAttachmentContentAsync(att.id, function (res) {
            if (
              res && res.status === Office.AsyncResultStatus.Succeeded && res.value &&
              res.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64
            ) {
              files.push({ name: att.name, contentType: att.contentType, base64: res.value.content });
            } else {
              lastErr = "1.8 status=" + (res && res.status) +
                (res && res.value ? " format=" + res.value.format : "");
            }
            resolve();
          });
        } catch (e) {
          lastErr = "1.8 exception: " + (e && e.message);
          resolve();
        }
      });
    });
  });
  return chain.then(function () {
    return { files: files, error: files.length ? "" : lastErr };
  });
}

// Доступен ли путь через Outlook REST API.
function canUseRest(item) {
  try {
    return !!(
      item.itemId &&
      Office.context.mailbox.getCallbackTokenAsync &&
      Office.context.mailbox.convertToRestId
    );
  } catch (e) {
    return false;
  }
}

// Способ 2: собрать файлы через Outlook REST API (/messages/{id}/attachments).
// Возвращает { files:[{name,contentType,base64}], error }.
function collectViaRest(item) {
  return new Promise(function (resolve) {
    try {
      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (tr) {
        if (!tr || tr.status !== Office.AsyncResultStatus.Succeeded) {
          resolve({ error: "REST токен: " + (tr && tr.error && tr.error.message) });
          return;
        }
        var token = tr.value;
        var restUrl = Office.context.mailbox.restUrl || "https://outlook.office365.com";
        var restId = Office.context.mailbox.convertToRestId(
          item.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
        var url = restUrl + "/api/v2.0/me/messages/" + encodeURIComponent(restId) + "/attachments";

        fetch(url, {
          headers: { Authorization: "Bearer " + token, Accept: "application/json" },
        })
          .then(function (resp) {
            return resp.text().then(function (t) {
              return { ok: resp.ok, status: resp.status, text: t };
            });
          })
          .then(function (r) {
            if (!r.ok) {
              resolve({ error: "REST " + r.status + ": " + (r.text || "").substring(0, 140) });
              return;
            }
            var data;
            try {
              data = JSON.parse(r.text);
            } catch (e) {
              resolve({ error: "REST parse error" });
              return;
            }
            var arr = data.value || [];
            var files = [];
            for (var i = 0; i < arr.length; i++) {
              var a = arr[i];
              // fileAttachment содержит ContentBytes (base64). itemAttachment пропускаем.
              if (a.ContentBytes) {
                files.push({ name: a.Name, contentType: a.ContentType, base64: a.ContentBytes });
              }
            }
            resolve({ files: files });
          })
          .catch(function (e) {
            resolve({ error: "REST fetch: " + ((e && e.message) || e) });
          });
      });
    } catch (e) {
      resolve({ error: "REST exception: " + ((e && e.message) || e) });
    }
  });
}

// Доступен ли путь через EWS (makeEwsRequestAsync + есть itemId).
function canUseEws(item) {
  try {
    return !!(item.itemId && Office.context.mailbox.makeEwsRequestAsync);
  } catch (e) {
    return false;
  }
}

// Способ 3: получить письмо целиком в MIME (EWS GetItem + IncludeMimeContent)
// и разобрать вложения на месте. GetItem — разрешённая для надстроек операция,
// в отличие от GetAttachment. Возвращает { files:[{name,contentType,base64}], error }.
function collectViaMime(item, atts) {
  return new Promise(function (resolve) {
    try {
      var soap = buildGetItemMimeSoap(item.itemId);
      Office.context.mailbox.makeEwsRequestAsync(soap, function (res) {
        if (!res || res.status !== Office.AsyncResultStatus.Succeeded) {
          resolve({ error: "MIME status=" + (res && res.status) + " " + (res && res.error && res.error.message) });
          return;
        }
        var body = res.value || "";
        var mm = /<(?:\w+:)?MimeContent[^>]*>([\s\S]*?)<\/(?:\w+:)?MimeContent>/.exec(body);
        if (!mm) {
          var codeM = /<(?:\w+:)?ResponseCode>([\s\S]*?)<\//.exec(body);
          var faultM = /<(?:\w+:)?(?:faultstring|MessageText)[^>]*>([\s\S]*?)<\//.exec(body);
          var info = codeM ? codeM[1] : (faultM ? faultM[1] : body.substring(0, 140));
          resolve({ error: "MIME без содержимого: " + info });
          return;
        }
        var raw;
        try {
          raw = window.atob(mm[1].replace(/\s+/g, ""));
        } catch (e) {
          resolve({ error: "MIME decode: " + (e && e.message) });
          return;
        }
        var files = parseMimeAttachments(raw, atts);
        resolve({ files: files });
      });
    } catch (e) {
      resolve({ error: "MIME exception: " + (e && e.message) });
    }
  });
}

function buildGetItemMimeSoap(itemId) {
  var id = String(itemId || "").replace(/[&<>"]/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
  });
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    ' xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    ' xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    ' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '<soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>' +
    "<soap:Body>" +
    '<GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
    ' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    "<ItemShape><t:BaseShape>IdOnly</t:BaseShape>" +
    "<t:IncludeMimeContent>true</t:IncludeMimeContent></ItemShape>" +
    '<ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    "</GetItem>" +
    "</soap:Body></soap:Envelope>"
  );
}

// Разбор MIME-письма: находим leaf-части с Content-Transfer-Encoding: base64
// и именем файла. Имя/тип по возможности берём из списка вложений по размеру.
function parseMimeAttachments(raw, atts) {
  var boundaries = {};
  var reB = /boundary="?([^";\r\n]+)"?/gi;
  var mb;
  while ((mb = reB.exec(raw))) {
    boundaries[mb[1]] = true;
  }

  var found = [];
  var seen = {};
  for (var b in boundaries) {
    if (!boundaries.hasOwnProperty(b)) continue;
    var chunks = raw.split("--" + b);
    for (var i = 0; i < chunks.length; i++) {
      var chunk = chunks[i];
      var nl = 4;
      var sep = chunk.indexOf("\r\n\r\n");
      if (sep < 0) { sep = chunk.indexOf("\n\n"); nl = 2; }
      if (sep < 0) continue;
      var headers = chunk.substring(0, sep);

      // Пропускаем вложенные multipart-контейнеры (это не сам файл).
      if (/content-type:\s*multipart\//i.test(headers)) continue;

      var fname = extractMimeFilename(headers);
      var isAttach = /content-disposition:\s*attachment/i.test(headers);
      // Берём только части-файлы: с именем файла или явным attachment.
      if (!fname && !isAttach) continue;

      // Тело части — до следующего boundary; отрезаем возможный хвост "--".
      var rawBody = chunk.substring(sep + nl).replace(/\r?\n--\s*$/, "").replace(/--\s*$/, "");
      var cte = extractCte(headers);
      var b64 = bodyToBase64(rawBody, cte);
      if (!b64 || b64.length < 4) continue;

      var sig = (fname || "") + ":" + b64.length + ":" + b64.substring(0, 16);
      if (seen[sig]) continue;
      seen[sig] = true;

      found.push({
        name: fname,
        contentType: extractMimeContentType(headers),
        base64: b64,
      });
    }
  }

  // Сопоставим имена/типы с реальными вложениями по размеру (декодированному).
  var used = {};
  for (var k = 0; k < found.length; k++) {
    var f = found[k];
    var approxLen = Math.floor((f.base64.length * 3) / 4);
    var best = -1;
    var bestDiff = 64; // допуск в байтах на паддинг/переносы
    for (var j = 0; j < atts.length; j++) {
      if (used[j] || typeof atts[j].size !== "number") continue;
      var diff = Math.abs(atts[j].size - approxLen);
      if (diff < bestDiff) { bestDiff = diff; best = j; }
    }
    if (best >= 0) {
      used[best] = true;
      f.name = atts[best].name || f.name;
      f.contentType = atts[best].contentType || f.contentType;
    }
    if (!f.name) f.name = "attachment_" + (k + 1);
  }
  return found;
}

function extractMimeFilename(headers) {
  var m =
    /filename\*?=\s*"?([^";\r\n]+)"?/i.exec(headers) ||
    /name\*?=\s*"?([^";\r\n]+)"?/i.exec(headers);
  if (!m) return "";
  return decodeMimeWord(m[1].replace(/"$/, ""));
}

function extractMimeContentType(headers) {
  var m = /content-type:\s*([^;\r\n]+)/i.exec(headers);
  return m ? m[1].trim() : "";
}

function extractCte(headers) {
  var m = /content-transfer-encoding:\s*([^\s;\r\n]+)/i.exec(headers);
  return m ? m[1].toLowerCase() : "";
}

// Приводит тело MIME-части к base64 в зависимости от Content-Transfer-Encoding.
// base64 → как есть; quoted-printable → декодируем; 7bit/8bit/binary → тело уже байты.
function bodyToBase64(body, cte) {
  cte = cte || "";
  if (cte === "base64") {
    return body.replace(/[^A-Za-z0-9+/=]/g, "");
  }
  var binary = cte === "quoted-printable" ? decodeQuotedPrintable(body) : body;
  return binaryToBase64(binary);
}

function decodeQuotedPrintable(s) {
  // Мягкие переносы строк ("=" в конце строки) убираем.
  s = s.replace(/=\r?\n/g, "");
  return s.replace(/=([0-9A-Fa-f]{2})/g, function (a, h) {
    return String.fromCharCode(parseInt(h, 16));
  });
}

// Байтовая строка (символы 0–255) → base64. Чанк кратен 3, чтобы склейка была корректной.
function binaryToBase64(binary) {
  try {
    if (binary.length <= 49152) return window.btoa(binary);
    var out = "";
    for (var i = 0; i < binary.length; i += 49152) {
      out += window.btoa(binary.substring(i, i + 49152));
    }
    return out;
  } catch (e) {
    return "";
  }
}

// Базовый разбор RFC2047 =?charset?B?..?= / =?charset?Q?..?= (best effort).
function decodeMimeWord(s) {
  try {
    return s.replace(/=\?[^?]+\?([BbQq])\?([^?]*)\?=/g, function (all, enc, txt) {
      if (enc.toUpperCase() === "B") {
        return utf8Decode(window.atob(txt));
      }
      var q = txt.replace(/_/g, " ").replace(/=([0-9A-Fa-f]{2})/g, function (a, h) {
        return String.fromCharCode(parseInt(h, 16));
      });
      return utf8Decode(q);
    });
  } catch (e) {
    return s;
  }
}

// Раскодировать байтовую строку как UTF-8 (IE11-совместимо).
function utf8Decode(bytes) {
  try {
    return decodeURIComponent(escape(bytes));
  } catch (e) {
    return bytes;
  }
}

// Base64 → Blob (IE11 поддерживает atob, Uint8Array и Blob).
function base64ToBlob(b64, contentType) {
  var binary = window.atob(b64);
  var len = binary.length;
  var bytes = new Uint8Array(len);
  for (var i = 0; i < len; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return new Blob([bytes], { type: contentType || "application/octet-stream" });
}

// Имя кастомного поля Kaiten, куда пишем отправителя. Можно переопределить
// в config.js через CUSTOMER_FIELD_NAME, по умолчанию — «Заказчик».
function customerFieldName() {
  return (window.KAITEN_CONFIG && window.KAITEN_CONFIG.CUSTOMER_FIELD_NAME) || "Заказчик";
}

// Ответ /company/custom-properties может прийти массивом или объектом с .data.
function normalizeProps(props) {
  if (Array.isArray(props)) return props;
  if (props && Array.isArray(props.data)) return props.data;
  return [];
}

// Находит id кастомного свойства по его названию (без учёта регистра/пробелов).
function findPropertyId(list, name) {
  var target = String(name || "").toLowerCase().replace(/^\s+|\s+$/g, "");
  for (var i = 0; i < list.length; i++) {
    var n = String(list[i].name || "").toLowerCase().replace(/^\s+|\s+$/g, "");
    if (n === target) return list[i].id;
  }
  return null;
}

// Строка отправителя письма: «Имя <email>», либо что есть.
function formatSender(data) {
  if (data.senderName && data.senderEmail) {
    return data.senderName + " <" + data.senderEmail + ">";
  }
  return data.senderName || data.senderEmail || "";
}

// Открывает URL карточки в браузере по умолчанию.
// Веб-надстройка не может выбрать конкретный браузер (Chrome и т.п.) — открывается
// системный браузер по умолчанию. openBrowserWindow предпочтительнее window.open.
function openInBrowser(url) {
  try {
    if (Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(url);
      return;
    }
  } catch (e) {
    /* fallback ниже */
  }
  try {
    window.open(url, "_blank");
  } catch (e2) {
    /* noop */
  }
}

function truncate(str, max) {
  if (!str) return "";
  if (str.length <= max) return str;
  return str.slice(0, max) + "\n\n[…обрезано, исходный текст в письме]";
}

function truncateLine(s, max) {
  if (!s) return "";
  return s.length <= max ? s : s.slice(0, max - 1) + "…";
}

function short(s) {
  return String(s || "").slice(0, 150);
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

function pad2(n) {
  n = String(n);
  return n.length < 2 ? "0" + n : n;
}

function formatDate(d) {
  return (
    pad2(d.getDate()) +
    "." +
    pad2(d.getMonth() + 1) +
    "." +
    d.getFullYear() +
    " " +
    pad2(d.getHours()) +
    ":" +
    pad2(d.getMinutes())
  );
}

function notify(item, key, message, type) {
  if (!item || !item.notificationMessages) return;
  try {
    item.notificationMessages.replaceAsync(key, {
      type: type,
      message: short(message),
      icon: "icon16",
      persistent: false,
    });
  } catch (e) {
    /* старые сборки Outlook 2016 могут не поддерживать notificationMessages */
  }
}

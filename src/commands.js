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
        notify(
          item,
          "kaiten-info",
          "Задача создана на доске «" + (choice.boardTitle || "?") + "»: #" + card.id,
          "informationalMessage"
        );
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

    var url =
      dialogUrl() +
      "?defBoard=" +
      encodeURIComponent(defaultBoardId) +
      "&defTitle=" +
      encodeURIComponent(defaultBoardTitle || "");

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

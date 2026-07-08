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
 */
function createKaitenCard(event) {
  var item = Office.context.mailbox.item;

  notify(item, "kaiten-info", "Создаю задачу в Kaiten…", "informationalMessage");

  run(item)
    .then(function (card) {
      var url = window.KaitenApi.getCardUrl(card.id);
      notify(item, "kaiten-info", "Задача создана: #" + card.id, "informationalMessage");
      if (window.KAITEN_CONFIG.OPEN_CARD_AFTER_CREATE) {
        try {
          window.open(url, "_blank");
        } catch (e) {
          /* noop */
        }
      }
    })
    .catch(function (err) {
      if (console && console.error) console.error("[Kaiten Add-in]", err);
      notify(item, "kaiten-error", short(err && err.message ? err.message : String(err)), "errorMessage");
    })
    .then(function () {
      // ExecuteFunction-команды обязаны вызывать event.completed().
      event.completed();
    });
}

/**
 * Основной сценарий создания карточки. Возвращает Promise с карточкой.
 */
function run(item) {
  var token = window.KaitenSettings.getToken();
  if (!token) {
    return Promise.reject(
      new Error("Не настроен токен. Открой «Настройки» в группе Kaiten на ленте.")
    );
  }
  var boardId = window.KaitenSettings.getBoardId();
  if (!boardId) {
    return Promise.reject(
      new Error("Не выбрана доска. Открой «Настройки» в группе Kaiten и выбери доску.")
    );
  }

  var card;

  return collectEmailData(item)
    .then(function (data) {
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

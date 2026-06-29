/* global Office, window, console */

Office.onReady(() => {
  if (typeof Office.actions !== "undefined" && Office.actions.associate) {
    Office.actions.associate("createKaitenCard", createKaitenCard);
  }
});

/**
 * Точка входа кнопки «Создать задачу» в группе Kaiten.
 */
function createKaitenCard(event) {
  const item = Office.context.mailbox.item;

  notify(item, "kaiten-info", "Создаю задачу в Kaiten…", "informationalMessage");

  run(item)
    .then((card) => {
      const url = window.KaitenApi.getCardUrl(card.id);
      notify(item, "kaiten-info", "Задача создана: #" + card.id, "informationalMessage");
      if (window.KAITEN_CONFIG.OPEN_CARD_AFTER_CREATE) {
        try {
          window.open(url, "_blank");
        } catch (e) {
          /* noop */
        }
      }
    })
    .catch((err) => {
      console.error("[Kaiten Add-in]", err);
      notify(item, "kaiten-error", short(err && err.message ? err.message : String(err)), "errorMessage");
    })
    .then(() => {
      // ExecuteFunction-команды обязаны вызывать event.completed().
      event.completed();
    });
}

async function run(item) {
  // Проверяем наличие токена и выбранной доски.
  const token = window.KaitenSettings.getToken();
  if (!token) {
    throw new Error(
      "Не настроен токен. Открой «Настройки» в группе Kaiten на ленте."
    );
  }
  const boardId = window.KaitenSettings.getBoardId();
  if (!boardId) {
    throw new Error(
      "Не выбрана доска. Открой «Настройки» в группе Kaiten и выбери доску."
    );
  }

  const data = await collectEmailData(item);
  const description = buildDescription(data, window.KAITEN_CONFIG);

  const payload = {
    board_id: boardId,
    title: truncateLine(data.subject, 250),
    description: description,
  };

  const columnId = window.KaitenSettings.getDefaultColumnId();
  if (columnId) payload.column_id = columnId;

  const card = await window.KaitenApi.createCard(payload);

  // Прикрепляем ссылку на исходное письмо как external link, если получится получить.
  try {
    if (item.itemId && Office.context.mailbox.convertToRestId) {
      const restId = Office.context.mailbox.convertToRestId(
        item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
      );
      const link = "https://outlook.office.com/owa/?ItemID=" + encodeURIComponent(restId) + "&exvsurl=1&viewmodel=ReadMessageItem";
      await window.KaitenApi.addExternalLink(card.id, link, "Исходное письмо в Outlook");
    }
  } catch (e) {
    // External link — best effort, провал не критичен.
    console.warn("[Kaiten Add-in] Не удалось добавить ссылку на письмо:", e);
  }

  return card;
}

/**
 * Собирает данные из открытого письма (тело — асинхронно).
 */
function collectEmailData(item) {
  return new Promise((resolve, reject) => {
    const data = {
      subject: item.subject || "(без темы)",
      senderName: "",
      senderEmail: "",
      body: "",
      attachments: [],
      receivedDate: item.dateTimeCreated ? new Date(item.dateTimeCreated) : null,
    };

    const sender = item.from || item.sender;
    if (sender) {
      data.senderName = sender.displayName || "";
      data.senderEmail = sender.emailAddress || "";
    }

    if (Array.isArray(item.attachments)) {
      data.attachments = item.attachments.map((a) => ({
        name: a.name || "(без имени)",
        size: typeof a.size === "number" ? a.size : null,
      }));
    }

    item.body.getAsync(Office.CoercionType.Text, (result) => {
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
  const lines = [];

  if (cfg.INCLUDE_SENDER_IN_DESCRIPTION && (data.senderName || data.senderEmail)) {
    const sender = data.senderName
      ? data.senderName + (data.senderEmail ? " <" + data.senderEmail + ">" : "")
      : data.senderEmail;
    lines.push("**От:** " + sender);
  }

  if (cfg.INCLUDE_RECEIVED_DATE_IN_DESCRIPTION && data.receivedDate) {
    lines.push("**Дата:** " + formatDate(data.receivedDate));
  }

  if (cfg.INCLUDE_ATTACHMENTS_IN_DESCRIPTION && data.attachments.length > 0) {
    lines.push("**Вложения:**");
    data.attachments.forEach((a) => {
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

function formatDate(d) {
  const pad = (n) => String(n).padStart(2, "0");
  return (
    pad(d.getDate()) +
    "." +
    pad(d.getMonth() + 1) +
    "." +
    d.getFullYear() +
    " " +
    pad(d.getHours()) +
    ":" +
    pad(d.getMinutes())
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

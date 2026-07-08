/* global window, fetch */

// Тонкий клиент для Kaiten REST API.
// Использует Bearer-токен из roamingSettings и BASE_URL из настроек/config.js.
//
// ВАЖНО: ES5-совместимый код (без async/await и стрелочных функций) — Outlook 2016
// работает на движке IE11. fetch/Promise подтягиваются полифилами в HTML.
window.KaitenApi = (function () {
  function getBaseUrl() {
    var fromSettings = window.KaitenSettings && window.KaitenSettings.getBaseUrl();
    if (fromSettings) return fromSettings;
    var cfg = window.KAITEN_CONFIG;
    if (cfg && cfg.BASE_URL) return cfg.BASE_URL.replace(/\/+$/, "");
    throw new Error(
      "Адрес Kaiten не задан. Открой «Настройки» в группе Kaiten и заполни поле «Адрес Kaiten»."
    );
  }

  function getSpaceId() {
    var fromSettings = window.KaitenSettings && window.KaitenSettings.getSpaceId();
    if (fromSettings) return fromSettings;
    return (window.KAITEN_CONFIG && window.KAITEN_CONFIG.SPACE_ID) || 0;
  }

  function getToken() {
    // Токен лежит в Office.context.roamingSettings — это per-user storage в Exchange-профиле.
    // Settings.js отвечает за его get/set, здесь только читаем.
    if (!window.KaitenSettings || !window.KaitenSettings.getToken) {
      throw new Error("KaitenSettings не инициализирован");
    }
    var token = window.KaitenSettings.getToken();
    if (!token) {
      throw new Error(
        "API-токен Kaiten не настроен. Открой Outlook → письмо → лента Kaiten → Настройки и вставь токен."
      );
    }
    return token;
  }

  /**
   * Универсальный запрос к API с Bearer-аутентификацией.
   * Возвращает Promise с распарсенным JSON (или null для 204).
   */
  function request(method, path, body) {
    var url, token, init;
    try {
      url = getBaseUrl() + path;
      token = getToken();
    } catch (e) {
      return Promise.reject(e);
    }

    init = {
      method: method,
      headers: {
        Accept: "application/json",
        Authorization: "Bearer " + token,
      },
    };

    if (body !== undefined) {
      init.headers["Content-Type"] = "application/json";
      init.body = JSON.stringify(body);
    }

    return fetch(url, init).then(function (resp) {
      if (!resp.ok) {
        // Пытаемся достать текст ошибки, потом бросаем.
        return resp.text().then(function (text) {
          var detail = "";
          if (text) {
            try {
              var err = JSON.parse(text);
              detail = err && err.message ? err.message : text;
            } catch (e) {
              detail = text;
            }
          }
          var msg =
            "Kaiten API " + resp.status + " " + resp.statusText + (detail ? ": " + detail : "");
          var error = new Error(msg);
          error.status = resp.status;
          throw error;
        });
      }

      if (resp.status === 204) return null;
      return resp.json();
    });
  }

  /**
   * Список досок в пространстве.
   * GET /api/latest/spaces/{spaceId}/boards
   */
  function listBoards(spaceId) {
    var id = spaceId || getSpaceId();
    if (!id) {
      return Promise.reject(
        new Error("ID пространства не задан. Открой «Настройки» и заполни поле «ID пространства».")
      );
    }
    return request("GET", "/api/latest/spaces/" + encodeURIComponent(id) + "/boards");
  }

  /**
   * Информация о текущем пользователе — используем как health-check токена.
   * GET /api/latest/users/current
   */
  function whoami() {
    return request("GET", "/api/latest/users/current");
  }

  /**
   * Создать карточку.
   * POST /api/latest/cards
   * Минимально нужен board_id и title.
   */
  function createCard(payload) {
    return request("POST", "/api/latest/cards", payload);
  }

  /**
   * Прикрепить готовый внешний URL к карточке (например, ссылку на письмо в Outlook).
   * POST /api/latest/cards/{cardId}/external-links
   */
  function addExternalLink(cardId, url, description) {
    return request("POST", "/api/latest/cards/" + cardId + "/external-links", {
      url: url,
      description: description || "",
    });
  }

  /**
   * URL карточки в браузере.
   */
  function getCardUrl(cardId) {
    return getBaseUrl() + "/space/" + getSpaceId() + "/card/" + cardId;
  }

  return {
    listBoards: listBoards,
    whoami: whoami,
    createCard: createCard,
    addExternalLink: addExternalLink,
    getCardUrl: getCardUrl,
  };
})();

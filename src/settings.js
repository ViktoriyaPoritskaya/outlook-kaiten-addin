/* global Office, window */

// Хранение пользовательских настроек (токен, выбранная доска) в roamingSettings.
// Это шифрованное хранилище в Exchange-профиле — данные не уезжают на хостинг
// и привязаны к конкретному пользователю.
//
// ВАЖНО: код намеренно написан на ES5 (var, function, без стрелочных функций
// и async/await), потому что Outlook 2016 desktop выполняет надстройки в движке
// Internet Explorer 11, который не понимает современный синтаксис.
window.KaitenSettings = (function () {
  var KEY_TOKEN = "kaiten.token";
  var KEY_BOARD_ID = "kaiten.boardId";
  var KEY_BOARD_TITLE = "kaiten.boardTitle";
  var KEY_DEFAULT_COLUMN_ID = "kaiten.columnId";
  var KEY_BASE_URL = "kaiten.baseUrl";
  var KEY_SPACE_ID = "kaiten.spaceId";

  function s() {
    if (
      !Office ||
      !Office.context ||
      !Office.context.roamingSettings
    ) {
      throw new Error("Office.context.roamingSettings недоступен в этом контексте");
    }
    return Office.context.roamingSettings;
  }

  function get(key) {
    try {
      var v = s().get(key);
      return v == null ? "" : v;
    } catch (e) {
      return "";
    }
  }

  function set(key, value) {
    s().set(key, value);
  }

  function save() {
    return new Promise(function (resolve, reject) {
      s().saveAsync(function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(res.error && res.error.message));
      });
    });
  }

  function trimTrailingSlash(v) {
    return (v || "").replace(/^\s+|\s+$/g, "").replace(/\/+$/, "");
  }

  return {
    getToken: function () { return get(KEY_TOKEN); },
    setToken: function (v) { return set(KEY_TOKEN, v); },
    getBoardId: function () { return Number(get(KEY_BOARD_ID)) || 0; },
    setBoardId: function (v) { return set(KEY_BOARD_ID, Number(v) || 0); },
    getBoardTitle: function () { return get(KEY_BOARD_TITLE); },
    setBoardTitle: function (v) { return set(KEY_BOARD_TITLE, v || ""); },
    getDefaultColumnId: function () { return Number(get(KEY_DEFAULT_COLUMN_ID)) || 0; },
    setDefaultColumnId: function (v) { return set(KEY_DEFAULT_COLUMN_ID, Number(v) || 0); },
    getBaseUrl: function () { return get(KEY_BASE_URL); },
    setBaseUrl: function (v) { return set(KEY_BASE_URL, trimTrailingSlash(v)); },
    getSpaceId: function () { return Number(get(KEY_SPACE_ID)) || 0; },
    setSpaceId: function (v) { return set(KEY_SPACE_ID, Number(v) || 0); },
    save: save,
  };
})();

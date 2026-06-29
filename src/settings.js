/* global Office, window */

// Хранение пользовательских настроек (токен, выбранная доска) в roamingSettings.
// Это шифрованное хранилище в Exchange-профиле — данные не уезжают на хостинг
// и привязаны к конкретному пользователю.
window.KaitenSettings = (function () {
  const KEY_TOKEN = "kaiten.token";
  const KEY_BOARD_ID = "kaiten.boardId";
  const KEY_BOARD_TITLE = "kaiten.boardTitle";
  const KEY_DEFAULT_COLUMN_ID = "kaiten.columnId";
  const KEY_BASE_URL = "kaiten.baseUrl";
  const KEY_SPACE_ID = "kaiten.spaceId";

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
      const v = s().get(key);
      return v == null ? "" : v;
    } catch (e) {
      return "";
    }
  }

  function set(key, value) {
    s().set(key, value);
  }

  function save() {
    return new Promise((resolve, reject) => {
      s().saveAsync((res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(res.error && res.error.message));
      });
    });
  }

  return {
    getToken: () => get(KEY_TOKEN),
    setToken: (v) => set(KEY_TOKEN, v),
    getBoardId: () => Number(get(KEY_BOARD_ID)) || 0,
    setBoardId: (v) => set(KEY_BOARD_ID, Number(v) || 0),
    getBoardTitle: () => get(KEY_BOARD_TITLE),
    setBoardTitle: (v) => set(KEY_BOARD_TITLE, v || ""),
    getDefaultColumnId: () => Number(get(KEY_DEFAULT_COLUMN_ID)) || 0,
    setDefaultColumnId: (v) => set(KEY_DEFAULT_COLUMN_ID, Number(v) || 0),
    getBaseUrl: () => get(KEY_BASE_URL),
    setBaseUrl: (v) => set(KEY_BASE_URL, (v || "").trim().replace(/\/+$/, "")),
    getSpaceId: () => Number(get(KEY_SPACE_ID)) || 0,
    setSpaceId: (v) => set(KEY_SPACE_ID, Number(v) || 0),
    save: save,
  };
})();

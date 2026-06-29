/* global window */

// Конфигурация расширения Outlook → Kaiten.
// Меняй только BASE_URL и SPACE_ID. Остальное — в настройках расширения
// (кнопка «Настройки» в группе Kaiten на ленте Outlook).
window.KAITEN_CONFIG = {
  // Адрес Kaiten и ID пространства задаются через кнопку «Настройки» в Outlook
  // и хранятся в зашифрованном профиле Exchange. В этом файле их намеренно нет.
  BASE_URL: "",
  SPACE_ID: 0,

  // Сколько символов тела письма максимум положить в описание карточки.
  // У Kaiten нет жёсткого лимита, но 8000 символов — разумный максимум.
  MAX_BODY_LENGTH: 8000,

  // Включать ли в описание карточки блоки с метаданными письма.
  INCLUDE_SENDER_IN_DESCRIPTION: true,
  INCLUDE_RECEIVED_DATE_IN_DESCRIPTION: true,
  INCLUDE_ATTACHMENTS_IN_DESCRIPTION: true,

  // Открывать ли созданную карточку в браузере после успешного создания.
  OPEN_CARD_AFTER_CREATE: true,
};

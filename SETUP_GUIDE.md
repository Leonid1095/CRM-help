# Пошаговая инструкция по установке CRM-Помощник на Ubuntu сервер

## Шаг 1: Подготовка — Получение токена бота

1. Откройте Telegram на телефоне или компьютере
2. Найдите бота **@BotFather** (официальный бот для создания ботов)
3. Напишите ему `/start`
4. Напишите `/newbot`
5. BotFather спросит название бота, напишите: **CRM-Помощник**
6. BotFather спросит username (должен заканчиваться на `bot`), напишите: **crm_1c_support_bot** (или другой, если занят)
7. BotFather даст вам **токен** — длинная строка типа `1234567890:ABCdefGHIjklMNOpqrsTUVwxyz`
8. **Скопируйте этот токен** — он понадобится позже

## Шаг 2: Узнать свой Telegram ID (для админки)

1. В Telegram найдите бота **@userinfobot**
2. Напишите ему `/start`
3. Он пришлёт ваш ID (цифры, например `123456789`)
4. **Запишите этот ID** — он понадобится для доступа к админ-панели

## Шаг 3: Подключиться к серверу Ubuntu

Откройте терминал (или PuTTY на Windows) и подключитесь:

```bash
ssh ваш_пользователь@IP_адрес_сервера
```

Введите пароль, когда попросит.

## Шаг 4: Обновить систему

```bash
sudo apt update
```

```bash
sudo apt upgrade -y
```

## Шаг 5: Установить необходимые программы

```bash
sudo apt install -y python3 python3-pip python3-venv git
```

## Шаг 6: Создать папку для бота

```bash
sudo mkdir -p /opt/crm-support-bot
```

```bash
sudo chown $USER:$USER /opt/crm-support-bot
```

## Шаг 7: Перейти в папку

```bash
cd /opt/crm-support-bot
```

## Шаг 8: Скачать код бота с GitHub

```bash
git clone https://github.com/Leonid1095/CRM-help.git .
```

(точка в конце важна — она означает "скачать прямо в текущую папку")

## Шаг 9: Создать виртуальное окружение Python

```bash
python3 -m venv venv
```

## Шаг 10: Активировать виртуальное окружение

```bash
source venv/bin/activate
```

После этого в начале строки появится `(venv)` — это нормально.

## Шаг 11: Установить зависимости бота

```bash
pip install -r requirements.txt
```

Подождите, пока установятся библиотеки (30-60 секунд).

## Шаг 12: Настроить бота

Откройте файл настроек:

```bash
nano config.py
```

Найдите строку:
```python
BOT_TOKEN = "ВАШ_ТОКЕН_СЮДА"
```

**Сотрите** `ВАШ_ТОКЕН_СЮДА` и **вставьте** токен, который дал BotFather (из Шага 1).

Найдите строки:
```python
ADMIN_IDS = [
    # 123456789,
]
```

Удалите `#` и вместо `123456789` вставьте **ваш Telegram ID** (из Шага 2), чтобы получилось:
```python
ADMIN_IDS = [
    123456789,
]
```

**Сохранить и выйти:**
- Нажмите `Ctrl + O` (сохранить)
- Нажмите `Enter` (подтвердить имя файла)
- Нажмите `Ctrl + X` (выйти)

## Шаг 13: Проверить, что бот запускается

```bash
python3 bot.py
```

Должно появиться: `CRM-Помощник запущен...`

Откройте Telegram, найдите вашего бота по username (из Шага 1), напишите `/start`.
Если бот отвечает — всё работает!

**Остановить бота:** нажмите `Ctrl + C` в терминале.

## Шаг 14: Настроить автозапуск (чтобы бот работал всегда)

Создайте файл службы:

```bash
sudo nano /etc/systemd/system/crm-bot.service
```

Скопируйте и вставьте следующий текст целиком:

```ini
[Unit]
Description=CRM Support Telegram Bot
After=network.target

[Service]
Type=simple
User=www-data
WorkingDirectory=/opt/crm-support-bot
ExecStart=/opt/crm-support-bot/venv/bin/python3 bot.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

**Сохранить и выйти:**
- `Ctrl + O`, `Enter`, `Ctrl + X`

## Шаг 15: Дать права папке для пользователя www-data

```bash
sudo chown -R www-data:www-data /opt/crm-support-bot
```

## Шаг 16: Включить и запустить службу

Перезагрузить systemd:
```bash
sudo systemctl daemon-reload
```

Включить автозапуск при перезагрузке сервера:
```bash
sudo systemctl enable crm-bot
```

Запустить бота:
```bash
sudo systemctl start crm-bot
```

Проверить статус:
```bash
sudo systemctl status crm-bot
```

Должно быть написано **active (running)** зелёным цветом.

Нажмите `q`, чтобы выйти из просмотра статуса.

## Шаг 17: Проверить в Telegram

Откройте Telegram, найдите бота, напишите `/start` — он должен работать!

---

## Дополнительные команды для управления ботом

### Посмотреть логи в реальном времени:
```bash
sudo journalctl -u crm-bot -f
```
(Выйти: `Ctrl + C`)

### Остановить бота:
```bash
sudo systemctl stop crm-bot
```

### Перезапустить бота (после изменений в коде):
```bash
sudo systemctl restart crm-bot
```

### Посмотреть статус:
```bash
sudo systemctl status crm-bot
```

---

## Если нужно изменить список модулей или категорий ошибок

1. Подключитесь к серверу
2. Откройте файл:
   ```bash
   sudo nano /opt/crm-support-bot/config.py
   ```
3. Измените списки `MODULES` или `ERROR_CATEGORIES`
4. Сохраните (`Ctrl + O`, `Enter`, `Ctrl + X`)
5. Перезапустите бота:
   ```bash
   sudo systemctl restart crm-bot
   ```

---

## Как пользоваться админкой

В Telegram напишите боту команду `/admin`.

Доступно:
- **Выгрузить Excel** — скачать файл со всеми обращениями
- **Статистика** — сколько обращений, ошибок, предложений
- **Список пользователей** — кто зарегистрирован

---

## Где хранятся данные

- **Пользователи:** `/opt/crm-support-bot/data/users.json`
- **Обращения (Excel):** `/opt/crm-support-bot/data/crm_support_log.xlsx`

Эти файлы создаются автоматически при первом запуске бота.

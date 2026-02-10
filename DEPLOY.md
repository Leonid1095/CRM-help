# Развёртывание CRM-Помощник на Ubuntu

## 1. Регистрация бота

1. Откройте Telegram, найдите **@BotFather**
2. Отправьте `/newbot`
3. Имя бота: `CRM-Помощник`
4. Username бота: `crm_1c_support_bot` (или любой свободный, например `crm_help_1c_bot`)
5. Скопируйте токен

## 2. Подготовка сервера

```bash
sudo apt update && sudo apt install -y python3 python3-pip python3-venv
```

## 3. Загрузка файлов

Скопируйте файлы проекта на сервер (например, через scp):

```bash
scp -r . user@server:/opt/crm-support-bot/
```

## 4. Установка зависимостей

```bash
cd /opt/crm-support-bot
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## 5. Настройка

Отредактируйте `config.py`:

```bash
nano config.py
```

- Вставьте токен бота в `BOT_TOKEN`
- Добавьте свой Telegram ID в `ADMIN_IDS` (узнать ID: напишите @userinfobot в Telegram)
- При необходимости отредактируйте списки `MODULES` и `ERROR_CATEGORIES`

## 6. Тестовый запуск

```bash
source venv/bin/activate
python3 bot.py
```

## 7. Автозапуск через systemd

```bash
sudo nano /etc/systemd/system/crm-bot.service
```

Содержимое:

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

```bash
sudo systemctl daemon-reload
sudo systemctl enable crm-bot
sudo systemctl start crm-bot
sudo systemctl status crm-bot
```

## 8. Управление

```bash
sudo systemctl stop crm-bot      # остановить
sudo systemctl restart crm-bot   # перезапустить
sudo journalctl -u crm-bot -f    # логи в реальном времени
```

## 9. Использование

### Для сотрудников:
- Найти бота в Telegram по username
- Нажать /start
- Пройти регистрацию (ФИО + модуль)
- Использовать меню для отправки ошибок/предложений

### Для администраторов:
- Отправить `/admin` боту
- Доступны: выгрузка Excel, статистика, список пользователей

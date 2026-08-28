# Настройки: бэкап vault в git и Remote Control

Дата: 2026-08-28

Разбирались с двумя вещами: резервное копирование базы и почему сессии Claude Code на телефоне пишут «не могу достучаться до компьютера».

## 1. Бэкап vault в GitHub

**Было:** git-репозиторий в `D:\Obsidian` есть, но remote не задан — бэкапа вне ПК нет. Вся автоматика плагина obsidian-git выключена.

**Сделали:**
- Создан приватный репозиторий `https://github.com/murat01042014/obsidian-vault.git`, привязан как `origin`, ветка `master`, первый push прошёл.
- В плагине obsidian-git включены авто-commit/push/pull. Интервал сначала 10 мин, потом сам поставил commit-and-sync на 3 мин + «backup after file change».

**Проверка:** коммиты `vault backup: <дата время>` появляются на
https://github.com/murat01042014/obsidian-vault/commits/master
Плагин коммитит только когда Obsidian запущен.

## 2. Remote Control — «Can't reach your computer (computer_unreachable)»

**Причина:** это не баг настроек и не из-за установки приложения на телефон. Remote Control — живой процесс на ПК. Сессия становится offline через секунды после закрытия Claude Desktop / процесса. Старые сессии («Волт проверка») — мёртвые записи, их надо удалить и создать новые.

**Сделали:** в `C:\Users\User\.claude\settings.json` выставлен `remoteControlAtStartup: true` — новые сессии сами включают Remote Control. Применяется после полного перезапуска Claude Desktop.

**Как поднять рабочую сессию:**
1. В Claude Desktop открыть сессию Claude Code в папке `D:\Obsidian`.
2. Набрать `/remote-control`.
3. Отсканировать показанный QR телефоном (или открыть ссылку).
4. Не закрывать Claude Desktop, не давать ПК уснуть.

**Требования:** план Pro подходит. Если Remote Control будет падать с ошибкой про API endpoint — убрать переменную окружения `ANTHROPIC_BASE_URL`.

## Заметки
- Скелет базы (`HOME.md`, `status.md`) ещё с плейсхолдерами — заполнить под себя.

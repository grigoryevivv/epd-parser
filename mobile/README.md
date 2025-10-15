# ЕПД Парсер - Mobile (Android)

Мобильное приложение на Kivy для Android устройств.

## 📱 Установка APK

### Автоматическая сборка через GitHub Actions:

1. Перейдите: https://github.com/grigoryevivv/epd-parser/actions
2. Выберите последний успешный workflow
3. Скачайте **"epdparser-debug-apk"** из Artifacts
4. Распакуйте ZIP → установите APK на телефон

## 🔨 Ручная сборка (для разработчиков)

### Требования:
- Linux (или WSL на Windows)
- Python 3.8+
- Buildozer

### Сборка:
```bash
# Установите buildozer
pip install buildozer

# Соберите APK
buildozer android debug

# APK будет в: bin/*.apk
```

## 🎯 GitHub Actions

Автоматическая сборка APK при каждом push в ветку main.

**Workflow файл:** `.github/workflows/build-apk.yml`

**Как запустить вручную:**
1. GitHub → Actions → "Build Android APK"
2. Run workflow → Run workflow
3. Ждать ~40 минут
4. Скачать APK из Artifacts

## 📊 Возможности приложения

- Загрузка PDF из памяти телефона
- Автоматическое распознавание данных
- Выбор услуг галочками
- Переключатель страхования
- Автоматический расчёт итогов
- Экспорт в Excel (сохраняется в Downloads)

## 📋 Файлы

- **main.py** - главное приложение (Kivy)
- **buildozer.spec** - конфигурация для сборки APK
- **requirements.txt** - зависимости Python
- **.github/workflows/build-apk.yml** - GitHub Actions workflow

## ⚙️ Требования

- Android 5.0 (API 21) или новее
- ~50 MB свободного места
- Разрешения: READ/WRITE_EXTERNAL_STORAGE

## 🔧 Решение проблем

### APK не устанавливается
- Включите "Установка из неизвестных источников" в настройках Android
- Проверьте версию Android (минимум 5.0)

### Приложение вылетает
- Проверьте разрешения на доступ к файлам
- Убедитесь что PDF не защищён

### Сборка на GitHub падает
- Проверьте логи в Actions
- Убедитесь что все файлы загружены
- Проверьте syntax в buildozer.spec

[app]

# Название приложения
title = EPD Parser

# Имя пакета
package.name = epdparser

# Домен пакета (используется для идентификации в Google Play)
package.domain = org.epdparser

# Исходный код приложения
source.dir = .

# Главный файл приложения
source.include_exts = py,png,jpg,kv,atlas,txt,pdf

# Версия приложения
version = 1.0

# Требования (Python модули)
requirements = python3,kivy,PyPDF2,pandas,openpyxl,python-dateutil,pyjnius,plyer,numpy,et-xmlfile,pillow

# Поддерживаемые ориентации (landscape, portrait, all)
orientation = portrait

# Полноэкранный режим
fullscreen = 0

# Разрешения Android
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,INTERNET

# Минимальная версия Android API
android.api = 31

# Минимальная поддерживаемая версия Android
android.minapi = 21

# Android NDK версия
android.ndk = 25b

# Поддержка архитектур
android.archs = arm64-v8a,armeabi-v7a

# Иконка приложения (если есть)
#icon.filename = %(source.dir)s/data/icon.png

# Splash экран (если есть)
#presplash.filename = %(source.dir)s/data/presplash.png

# Цвет фона splash экрана
#android.presplash_color = #FFFFFF

[buildozer]

# Логирование (0 = только ошибки, 1 = info, 2 = debug)
log_level = 2

# Количество потоков для сборки
#parallel = 4

# Путь к Android SDK (если не в PATH)
#android.sdk_path =

# Путь к Android NDK (если не в PATH)
#android.ndk_path =

# Директория для временных файлов
#build_dir = ./.buildozer

# Директория для бинарных файлов
#bin_dir = ./bin

# Автоматическое принятие лицензий Android SDK
android.accept_sdk_license = True

# Очистка временных файлов после сборки
#warn_on_root = 1

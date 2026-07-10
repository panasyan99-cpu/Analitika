АНАЛИТИКА v1.0 — сборка Windows-приложения

1. Установить Python 3.11+ для сборки.
2. Распаковать проект.
3. Запустить build_exe.bat.
4. Готовая папка для сотрудников появится в release/Аналитика.

Сотрудникам отправлять именно release/Аналитика.
Python и PyCharm сотрудникам не нужны.

Если сборка падает:
- запустить build_exe_debug.bat;
- посмотреть текст ошибки;
- проверить, что установлены зависимости из requirements.txt.

Основные файлы:
- Analitika.py — интерфейс приложения.
- src/report.py — ядро формирования отчета.
- assets/logo.png — логотип в окне.
- assets/analitika.ico — иконка EXE/установщика.
- Analitika.spec — сборка PyInstaller.
- installer_inno_setup.iss — установщик Inno Setup.

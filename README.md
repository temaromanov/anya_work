📄 Автоматизация договоров аренды и актов
Простое настольное приложение для автоматического заполнения Word-договоров, актов и других документов по шаблонам — с вводом данных через удобный интерфейс или Excel.

🚀 Возможности
Заполнение договоров аренды и актов по шаблонам Word (.docx)
Поддержка шаблонов для ООО и ИП
Генерация актов возврата, передаточных актов, актов оказания услуг
Автоматическая обработка ФИО, дат, сумм, НДС, сумма прописью
Сохранение данных в Excel
Интуитивно понятный интерфейс (tkinter)
Все данные хранятся у пользователя — никакого облака

🛠️ Как установить:
Клонируй репозиторий:
git clone https://github.com/temaromanovv/Programm_arenda_dogovor.git
cd Programm_arenda_dogovor

Создай виртуальное окружение и активируй его:
python -m venv .venv
.venv\Scripts\activate   # Windows
# source .venv/bin/activate   # Linux/Mac

Установи зависимости:
pip install -r requirements.txt

Запусти приложение:
python main.py
или, если собран exe:
main.exe

🖨️ Как собрать exe самостоятельно:

Установи PyInstaller внутри виртуального окружения:
pip install pyinstaller

Собери exe:
python -m PyInstaller --onefile --windowed main.py
Готовый файл будет в папке dist/main.exe.

📁 Структура папки
/Programm_arenda_dogovor/
    main.py
    form_ooo.py
    form_ip.py
    uslugi_ooo.py
    uslugi_ip.py
    act_vozvrata_ooo.py
    peredatochn_act_ooo.py
    main.spec
    requirements.txt
    README.md
    Шаблон_аренда_договор_ООО.docx
    ...
main.exe (если нужен запуск без Python) — не хранится в репозитории!

Все Word/Excel-шаблоны должны лежать рядом с main.exe/main.py

⚠️ Важно
Не заливай в репозиторий dist/, build/, .venv/, .exe-файлы — они генерируются на каждом ПК отдельно.

Весь функционал работает только под Windows.

📬 Обратная связь
Автор: Артём
Почта: temaromanovv@gmail.com
Telegram: @atrem_ai

P.S. Спасибо за интерес к проекту!
Если есть вопросы — пиши issue или напрямую.

💡 Хочешь — укажи конкретные шаблоны, примеры использования или добавь скриншоты UI для наглядности!
Если нужны дополнительные инструкции (например, “Как добавить новый шаблон” или “Как доработать логику”) — пиши, добавлю!

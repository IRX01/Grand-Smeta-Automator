import time
import os
import pyautogui
import psutil
from openpyxl import load_workbook
from pywinauto import Application
import pyperclip
import win32com.client
import re
import calendar
from datetime import datetime

def run_bot(Root_Folder):
    ##
    ##
    #Функция сохранения базовой сметы
    ##
    ##
    def save_excel_to_base_smeta(root_folder):

        # подключаемся к Excel
        app = Application(backend="uia").connect(title_re=".*Excel.*")
        excel = app.window(title_re=".*Excel.*")

        title = excel.window_text()

        # убираем "- Excel"
        title = title.replace(" - Excel", "")

        # обрезаем всё после "-"
        base_name = title.split("-")[0].strip()
        smeta_number = title.split("№")[1].split(" ")[0]

        print("Имя сметы:", base_name)

        # ищем папку
        target_folder = None

        for root, dirs, files in os.walk(root_folder):
            if smeta_number in root:
                target_folder = root
                break

        if not target_folder:
            print("Папка со сметой не найдена")
            return

        print("Найдена папка:", target_folder)

        # формируем путь
        save_path = os.path.join(
            target_folder,
            base_name
        )

        print("Путь сохранения:", save_path)

        time.sleep(1)

        # активируем поле имени файла
        pyautogui.hotkey("alt", "n")

        time.sleep(0.5)

        # вставляем путь
        pyperclip.copy(save_path)
        pyautogui.hotkey("ctrl", "v")

        time.sleep(0.5)

        # нажимаем сохранить
        pyautogui.press("enter")

        print("Файл базовой сметы сохранён")
    ##
    ##
    ##Функция сохранения базовой сметы
    ##
    ##

    ##
    ##
    ##Функция редактирования ДА
    ##
    ##
    def edit_DA_excel():
    
        excel = win32com.client.GetActiveObject("Excel.Application")
        wb = excel.ActiveWorkbook
        ws = wb.ActiveSheet
    
        last_row = ws.UsedRange.Rows.Count
        last_col = ws.UsedRange.Columns.Count
    
        rows_to_delete = []
        engineer_row = None
        
    
        for row in range(1, last_row + 1):
    
            for col in range(1, last_col + 1):
    
                value = ws.Cells(row, col).Value
    
                if value is None:
                    continue
    
                text = str(value)
    
                # вставляем номер сметы
                if "ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ №" in text:
    
                    number = wb.Name.split("№")[1].split(" ")[0]
                    ws.Cells(row, col).Value = f"ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ №{number}"
    
                # ищем строку с инженером
                if "Главный инженер" in text:
                    engineer_row = row

                
    
                delete_patterns = [
                    "Дата составления сметы",
                    "Согласовано",
                    "Проверил",
                    "Сдал",
                    "Заказчик",
                    "М.П.",
                    "должность"
                ]
    
                if any(p in text for p in delete_patterns):
                    rows_to_delete.append(row)
        
    
        # оставляем строку с инженером и строку под ней
        if engineer_row:
            if engineer_row in rows_to_delete:
                rows_to_delete.remove(engineer_row)
    
            if engineer_row + 1 in rows_to_delete:
                rows_to_delete.remove(engineer_row + 1)
    
        # удаляем строки снизу вверх
        for r in sorted(set(rows_to_delete), reverse=True):
            ws.Rows(r).Delete()
    
        print("ДА отредактирован")

    ##
    ##
    ##Функция редактирования ДА
    ##
    ##
        
    ##
    ##
    ## Функция редактирования КС-2
    ##
    ##
    def edit_ks2_excel():
    
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except:
            excel = win32com.client.Dispatch("Excel.Application")
    
        wb = excel.ActiveWorkbook
        ws = wb.ActiveSheet
    
        last_row = ws.UsedRange.Rows.Count
        last_col = ws.UsedRange.Columns.Count
    
        smeta_row = None
        number_row = None
        smeta_sum = None
    
        # --- первый проход (стоимость + номер) ---
    
        for row in range(1, last_row + 1):
    
            row_text = ""
    
            for col in range(1, last_col + 1):
                value = ws.Cells(row, col).Value
                if value:
                    row_text += str(value) + " "
    
            row_text = row_text.strip()
    
            if "Сметная (договорная) стоимость" in row_text:
                smeta_row = row
    
                match = re.search(r"(\d+[.,]\d+)", row_text)
                if match:
                    pyperclip.copy(match.group(1))
                    smeta_sum = match.group(1)
                    print("Сумма:", match.group(1))
    
            if row_text.lower().startswith("номер"):
                number_row = row
    
        # --- удаляем между сметой и номером ---
    
        if smeta_row and number_row:
            for r in range(number_row - 2, smeta_row, -1):
                ws.Rows(r).Delete()
    
        print("Блок перед таблицей очищен")
    
        # --- после удаления заново ищем ВСЕГО и ИНЖЕНЕР ---
    
        last_row = ws.UsedRange.Rows.Count
    
        total_row = None
        engineer_row = None
    
        for row in range(1, last_row + 1):
    
            row_text = ""
    
            for col in range(1, last_col + 1):
                value = ws.Cells(row, col).Value
                if value:
                    row_text += str(value) + " "
    
            row_text = row_text.strip()
    
            if "ВСЕГО по акту" in row_text:
                total_row = row
    
            if "Главный инженер" in row_text:
                engineer_row = row
    
        # --- удаляем между ВСЕГО и ИНЖЕНЕР ---
    
        if total_row and engineer_row:
            for r in range(engineer_row - 2, total_row, -1):
                ws.Rows(r).Delete()
    
        print("Хвост документа очищен")

        return smeta_sum
    ##
    ##
    ## Функция редактирования КС-2
    ##
    ##
        
    ##
    ##
    ## Функция открытия КС-3
    ##
    ##
    def open_ks3(root_folder, base_name):
    
        ks3_file = None
    
        for root, dirs, files in os.walk(root_folder):
    
            for file in files:
    
                if file.startswith("КС-3") and base_name in file:
                    ks3_file = os.path.join(root, file)
                    break
    
        if not ks3_file:
            print("КС-3 не найден")
            return None
    
        print("Найден файл:", ks3_file)
    
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
    
        wb = excel.Workbooks.Open(ks3_file)
    
        return wb
    ##
    ##
    ## Функция открытия КС-3
    ##
    ##

    ##
    ##
    ## Функция редактирования КС-3
    ##
    ##
    def edit_ks3(wb, base_name, address, _first_day, _last_day, sum_value):

        ws = wb.ActiveSheet

        number = base_name.split("№")[1].split(" ")[0]

        first_day = _first_day
        last_day =_last_day

        

        # --- адрес стройки ---
        ws.Range("A10").Value = f"Стройка - МКД по адресу: г. Тула, ул. {address}"

        # --- таблица ---
        ws.Range("E18").Value = number
        ws.Range("F18").Value = last_day
        ws.Range("H18").Value = first_day
        ws.Range("I18").Value = last_day
        

        # --- сумма ---
        ws.Range("I25").Value = sum_value
        ws.Range("I34").Value = sum_value


        print("КС-3 отредактирован")
    ##
    ##
    ## Функция редактирования КС-3
    ##
    ##
        

    def get_date_from_estimate(estimate_name):
        
        match = re.search(r"\d{2}\.\d{2}\.\d{4}", estimate_name)

        if match:
            return match.group()
        return None

    def get_last_day_of_month(date_string):

        date_obj = datetime.strptime(date_string, "%d.%m.%Y")

        last_day = calendar.monthrange(date_obj.year, date_obj.month)[1]

        result = f"{last_day:02d}.{date_obj.month:02d}.{date_obj.year}"

        return result

    def get_first_day_of_month(date_string):
        
        Date_obj = datetime.strptime(date_string, "%d.%m.%Y")

        first_day = Date_obj.replace(day=1)
        
        first_day = first_day.strftime("%d.%m.%Y")

        return first_day
        

    ##
    ##
    ## Функция сохранения ДА
    ##
    ##
    def save_excel_DA(root_folder):

        # подключаемся к Excel
        app = Application(backend="uia").connect(title_re=".*Excel.*")
        excel = app.window(title_re=".*Excel.*")

        title = excel.window_text()

        # убираем "- Excel"
        title = title.replace(" - Excel", "")
        smeta_number = title.split("№")[1].split(" ")[0]

        # обрезаем всё после "-"
        base_name = title.split("-")[0].strip()

        print("Имя сметы:", base_name)

        # ищем папку
        target_folder = None

        for root, dirs, files in os.walk(root_folder):
            if smeta_number in root:
                target_folder = root
                break

        if not target_folder:
            print("Папка со сметой не найдена")
            return

        print("Найдена папка:", target_folder)

        # формируем путь
        save_path = os.path.join(
            target_folder,
            "ДА. " + base_name
        )

        print("Путь сохранения:", save_path)

        time.sleep(1)

        # активируем поле имени файла
        pyautogui.hotkey("alt", "n")

        time.sleep(0.5)

        # вставляем путь
        pyperclip.copy(save_path)
        pyautogui.hotkey("ctrl", "v")

        time.sleep(0.5)

        # нажимаем сохранить
        pyautogui.press("enter")

        print("ДА сохранён")
    ##
    ##
    ## Функция сохранения ДА
    ##
    ##
        
    ##
    ##
    ## Функция сохранения КС-2
    ##
    ##
    def save_excel_ks2(root_folder):

        # подключаемся к Excel
        app = Application(backend="uia").connect(title_re=".*Excel.*")
        excel = app.window(title_re=".*Excel.*")

        title = excel.window_text()

        # убираем "- Excel"
        title = title.replace(" - Excel", "")
        smeta_number = title.split("№")[1].split(" ")[0]

        # обрезаем всё после "-"
        base_name = title.split("-")[0].strip()

        print("Имя сметы:", base_name)

        # ищем папку
        target_folder = None

        for root, dirs, files in os.walk(root_folder):
            if smeta_number in root:
                target_folder = root
                break

        if not target_folder:
            print("Папка со сметой не найдена")
            return

        print("Найдена папка:", target_folder)

        # формируем путь
        save_path = os.path.join(
            target_folder,
            "КС-2. " + base_name
        )

        print("Путь сохранения:", save_path)

        time.sleep(1)

        # активируем поле имени файла
        pyautogui.hotkey("alt", "n")

        time.sleep(0.5)

        # вставляем путь
        pyperclip.copy(save_path)
        pyautogui.hotkey("ctrl", "v")

        time.sleep(0.5)

        # нажимаем сохранить
        pyautogui.press("enter")

        print("КС-2 сохранён")

##------------------------------------------------------------------------------
##------------------------------------------------------------------------------
## НАЧАЛО РАБОТЫ СКРИПТА
##------------------------------------------------------------------------------
##------------------------------------------------------------------------------
    print("Скрипт запущен")
    time.sleep(3)

    # активируем Excel
    #pyautogui.hotkey('alt','f')

    time.sleep(0.5)

    #pyautogui.click(x=100, y=400)
    pyautogui.doubleClick(x=250, y=350)
    pyautogui.click(x=830, y=640)

    while True:
        try:
            app = Application(backend="uia").connect(title_re=".*Excel*")
            excel = app.window(title_re=".*Excel*")
            excel.wait("visible", timeout=60)
            break
        except:
            time.sleep(1)

    

    ##
    ##Получение названия сметы, даты и номера для нового акта в КС-2
    ##
    App = Application(backend="uia").connect(title_re=".*Excel.*")
    Excel = App.window(title_re=".*Excel.*")
    Title = Excel.window_text()  
    Title = Title.replace(" - Excel", "")
    Base_name = Title.split("-")[0].strip()
    Adress = Base_name.split("г. ")[1]
    Estimate_number = Base_name.split(" ")[0]
    Estimate_number = Estimate_number.split("№")[1]
    print(f"НОМЕР СМЕТЫ БЛЯТЬ---- {Estimate_number}")
    date = get_date_from_estimate(Base_name)
    Last_day_date = get_last_day_of_month(date)
    First_day_date = get_first_day_of_month(date)
    print(f"ПОСЛЕДНИЙ ДЕНЬ БЛЯТЬ---- {Last_day_date}")
    print(f"ПЕРВЫЙ ДЕНЬ БЛЯТЬ---- {First_day_date}")
    print(f"АДРЕС БЛЯТЬ ---- {Adress}")
    ##
    ##Получение названия сметы, даты и номера для нового акта
    ##

    ##
    ##Сохранение базовой сметы
    ##
    excel.set_focus()
    time.sleep(0.5)
    pyautogui.hotkey('alt','f4')
    time.sleep(0.5)
    pyautogui.press("enter")
    ROOT_FOLDER = Root_Folder
    save_excel_to_base_smeta(ROOT_FOLDER)
    ##
    ##Сохранение базовой сметы
    ##

    time.sleep(1)
    pyautogui.click(x=250,y=400)
    pyautogui.doubleClick(x=250, y=400)

    while True:
        try:
            app = Application(backend="uia").connect(title_re=".*Excel*")
            excel = app.window(title_re=".*Excel*")
            excel.wait("visible", timeout=60)
            break
        except:
            time.sleep(1)

    ##
    ## Редактирование и сохранение ДА
    ##
    edit_DA_excel() 
    excel.set_focus()
    time.sleep(1)
    pyautogui.hotkey('alt','f4')
    time.sleep(1)
    pyautogui.press("enter")
    save_excel_DA(ROOT_FOLDER)
    time.sleep(1)
    ##
    ## Редактирование и сохранение ДА
    ##

    ##
    ## Создание нового акта в Гранд-смете
    ##
    pyautogui.click(x=605, y=55)
    time.sleep(0.5)
    pyautogui.click(x=130, y=130)
    time.sleep(0.5)
    pyautogui.click(x=405, y=230)
    time.sleep(0.5)
    pyautogui.click(x=740, y= 120)
    time.sleep(0.5)
    pyautogui.click(x=205, y= 125)
    time.sleep(0.5)
    pyautogui.click(x=490, y= 165)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl','a')
    pyautogui.press("backspace")
    pyautogui.typewrite(Estimate_number)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.press("backspace")
    pyautogui.typewrite(Last_day_date)
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.typewrite(Last_day_date)
    pyautogui.press("enter")
    ##
    ## Создание нового акта в Гранд-смете
    ##

    pyautogui.click(x = 40, y = 55)
    pyautogui.click(x = 100, y = 400)
    pyautogui.doubleClick(x = 625, y = 340)

    while True:
        try:
            app = Application(backend="uia").connect(title_re=".*Excel*")
            excel = app.window(title_re=".*Excel*")
            excel.wait("visible", timeout=60)
            break
        except:
            time.sleep(1)


    ##
    ## Редактирование и сохранение КС-2
    ##
    Sum_value = edit_ks2_excel()
    excel.set_focus()
    time.sleep(0.5)
    pyautogui.hotkey('alt','f4')
    time.sleep(0.5)
    pyautogui.press("enter")
    save_excel_ks2(ROOT_FOLDER)
    ##
    ## Редактирование и сохранение КС-2
    ##

    ##
    ## Поиск, открытие, редактирование и сохранение КС-3
    ##
    wb = open_ks3(ROOT_FOLDER, Base_name)
    while True:
        try:
            app = Application(backend="uia").connect(title_re=".*Excel*")
            excel = app.window(title_re=".*Excel*")
            excel.wait("visible", timeout=60)
            break
        except:
            time.sleep(1)
    edit_ks3(wb, Base_name, Adress, First_day_date, Last_day_date, Sum_value)
    excel.set_focus()
    time.sleep(0.5)
    pyautogui.hotkey('alt','f4')
    time.sleep(0.5)
    pyautogui.press("enter")
    ##
    ## Поиск, открытие, редактирование и сохранение КС-3
    ##




    print("Скрипт закончил свою работу")

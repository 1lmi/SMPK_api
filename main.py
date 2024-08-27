import time
import logging
from fastapi import FastAPI, HTTPException, Request
from datetime import datetime
from isoweek import Week
import openpyxl

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

groups = [
    'ТОР-23', 'РЭГ-23', 'СЭЗС-23', 'ПР-23', 'ОПИ-23', 'ДПИ-23', 'МД-23/1', 'МД-23/2',
    'ИСИП-23/1', 'ИСИП-23/2', 'БУ-23', 'БД-23', 'Ф-23', 'ЗИМ-23', 'ЮР-23/1', 'ЮР-23/2',
    'ПКД-23', 'ТОР-22', 'РЭГ-22', 'КИП-22', 'СЭЗС-22', 'ПР-22', 'ОПИ-22', 'ДПИ-22', 'МД-22',
    'ПО-22/1', 'ПО-22/2', 'БУ-22', 'БД-22', 'Ф-22', 'ЗИМ-22', 'ПСО-22/1', 'ПСО-22/2', 'ПКД-22',
    'ТОР-21', 'РЭГ-21', 'КИП-21', 'СЭЗС-21', 'ПР-21', 'ОПИ-21', 'ДПИ-21', 'МД-21', 'ПО-21',
    'БУ-21', 'Ф-21', 'ЗИМ-21', 'ПСО-21/1', 'ПСО-21/2', 'ПКД-21', 'ТОР-20', 'РЭГ-20', 'СЭЗС-20',
    'ПР-20', 'ОПИ-20', 'ДПИ-20', 'БУ-20', 'МД-20', 'ПО-20', 'ПКД-20'
]


def get_lesson_number(i):
    """Возвращает номер пары на основе строки в расписании."""
    lesson_numbers = ["1", "2", "3", "4", "5", "6"]
    return lesson_numbers[i]


def load_schedule_file(filename):
    """Загружает и возвращает активный лист из файла Excel."""
    try:
        wb_schedule = openpyxl.load_workbook(filename)
        return wb_schedule.active
    except FileNotFoundError:
        logger.error(f"Файл {filename} не найден.")
        raise HTTPException(status_code=404, detail="Schedule file not found")
    except Exception as e:
        logger.error(f"Ошибка при загрузке файла {filename}: {e}")
        raise HTTPException(status_code=500, detail="Internal Server Error")

def get_schedule_for_day(user_group: str, day: str):
    """Получает расписание на указанный день для заданной группы."""
    start_time = time.time()
    today = datetime.today()
    current_day = today.weekday()
    current_week = Week.withdate(today).week

    if user_group not in groups:
        raise HTTPException(status_code=404, detail="Group not found")

    schedule_column = groups.index(user_group) + 3

    schedule_filename = "rasp_cet.xlsx" if current_week % 2 == 0 else "rasp_necet.xlsx"

    if current_day == 6:  # Если сегодня воскресенье
        schedule_filename = "rasp_cet.xlsx" if current_week % 2 != 0 else "rasp_necet.xlsx"

    sheet_schedule = load_schedule_file(schedule_filename)

    if day == "tomorrow":
        target_day = (current_day + 1) % 7  # Завтрашний день
    elif day == "today":
        target_day = current_day  # Сегодня
    else:
        raise HTTPException(status_code=400, detail="Invalid day specified")

    start_row = 6 + (target_day * 6)
    end_row = start_row + 5

    schedule = {}
    for i in range(start_row, end_row + 1):
        lesson = sheet_schedule.cell(row=i, column=schedule_column).value
        cleaned_lesson = " ".join(lesson.split()) if lesson else "Нет"
        schedule[get_lesson_number(i - start_row)] = cleaned_lesson

    end_time = time.time()
    logger.info(f"Время выполнения запроса для группы {user_group} на {day}: {end_time - start_time:.4f} секунд")
    return schedule


@app.get("/schedule/")
async def read_schedule(day: str, group: str):
    # Выводим тело запроса в консоль
    print(f"Received request with parameters: day={day}, group={group}")
    return get_schedule_for_day(group, day)


@app.get("/")
async def read():
    return {"message": "Hello World"}

from PyPDF2 import PdfReader
from utils.employee_utils import Employee, TimeInterval

def delete_header_footer(page_text):
    index = page_text.find("Horario por centro de costo: Semanal (Excel)")
    return page_text[:index]


def get_pdf_text(pdf_path):
    reader = PdfReader(pdf_path)
    pdf_text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        page_text = delete_header_footer(page_text)
        pdf_text += page_text
    pdf_text = pdf_text[pdf_text.find("dom")+14:]
    return pdf_text


def get_pdf_list(pdf_text):
    pdf_text = pdf_text.replace("RS", " ")
    pdf_text = pdf_text.replace("SELF", " ")
    pdf_text = pdf_text.replace("CAJEROS", " ")
    pdf_text = pdf_text.replace("DIA DESCANSO\n0:00", " 0_DDD ")
    pdf_text = pdf_text.replace("DIA DESCANSO \n0:00", " 0_DDD ")
    pdf_text = pdf_text.replace("DIA DE\nDESCANSO 0:00", " 0_DDD ")
    pdf_text = pdf_text.replace("DIA DE \nDESCANSO 0:00", " 0_DDD ")
    pdf_text = pdf_text.replace("VACACIONES 0:00", " 0_VAC ")
    pdf_text = pdf_text.replace("PAGO HORAS\nFERIADO", " 0_PHF ")
    pdf_text = pdf_text.replace("PAGO HORAS \nFERIADO", " 0_PHF ")

    pdf_list = pdf_text.split()

    i = 0
    n = len(pdf_list)

    while i < n:
        if pdf_list[i] == "0_PHF" and pdf_list[i+1][0].isdigit():
            # Eliminar la duración de los PAGOS DE HORAS FERIADO
            holiday_interval_duration_index = pdf_list[i+1].find(":") + 3
            pdf_list[i+1] = pdf_list[i+1][holiday_interval_duration_index:]
            if pdf_list[i+1] == "":
                pdf_list.pop(i+1)
                n -= 1

        if (not pdf_list[i][0].isdigit() or pdf_list[i] == "x"):
            if not (i == 0 or (pdf_list[i-1][0].isdigit() or pdf_list[i-1] == "x")):
                pdf_list[i-1] += f" {pdf_list[i]}"
                pdf_list.pop(i)
                n -= 1
                if pdf_list[i-1][-2].isdigit():
                    pdf_list.insert(i, pdf_list[i-1][-13:])
                    pdf_list[i-1] = pdf_list[i-1][:-13]
                    n += 1
                continue

        i += 1

    return pdf_list


def format_schedule(schedule: str) -> TimeInterval:
    entrada = schedule.split("-")[0]
    salida = schedule.split("-")[1]

    am_format = "AM"
    pm_format = "PM"

    hora = int(entrada.split(":")[0])
    if (int(hora) >= 12):
        hora = "12"
        entrada = f"{hora}:{entrada.split(':')[1]}"

    if (entrada[-1] == "a"):
        entrada = entrada[:-1] + am_format
    else:
        entrada = entrada[:-1] + pm_format

    hora = int(salida.split(":")[0])
    if (int(hora) >= 12):
        hora = "12"
        salida = f"{hora}:{salida.split(':')[1]}"
    if (salida[-1] == "a"):
        salida = salida[:-1] + am_format
    else:
        salida = salida[:-1] + pm_format

    t = TimeInterval()
    t.set_interval(entrada, salida)

    return t


def format_item(itemString: str) -> str | TimeInterval:
    if itemString == "0_DDD":
        return "DIA DE DESCANSO"
    if itemString == "0_VAC":
        return "VACACIONES"
    if itemString == "0_PHF":
        return "PAGO HORAS FERIADO"
    return format_schedule(itemString)


def get_employees_list(pdf_list, self_amount=6):
    employees_list = []
    current_employee = Employee()
    self_counter = 1
    current_category_str = ""
    current_day_index = 0

    for item in pdf_list:
        if not (item[0].isdigit() or item == "x"):
            # Nombre
            if current_employee.name is None:
                current_employee.name = item
            else:
                current_employee.category = current_category_str
                if (current_employee.category == "SELF"):
                    self_counter += 1
                employees_list.append(current_employee)
                current_employee = Employee()
                current_employee.name = item
            current_category_str = "RS"
        elif (item == "x"):
            # Categoría
            if (self_counter > self_amount):
                current_category_str = "CAJERO"
            else:
                current_category_str = "SELF"
        else:
            formated_item = format_item(item)
            if (type(formated_item) == TimeInterval):
                current_employee.schedule.get_day_by_index(
                    current_day_index).interval = formated_item
                current_day_index += 1
                current_day_index = current_day_index % 7
            elif (type(formated_item) == str):
                current_employee.schedule.get_day_by_index(
                    current_day_index).day_type = formated_item
                current_day_index += 1
                current_day_index = current_day_index % 7

    return employees_list

def scrap_pdf(pdf_path, self_amount=6) -> list[Employee]:
    pdf_text = get_pdf_text(pdf_path)
    pdf_list = get_pdf_list(pdf_text)
    employees_list = get_employees_list(pdf_list, self_amount)
    return employees_list

import pandas as pd
from openpyxl import load_workbook, Workbook
import os

from utils.employee_utils import *
from utils.pdf_utils import scrap_pdf
from utils.time_utils import strfdelta
from datetime import timedelta, datetime
import clipboard
#from tkinter import filedialog
#from tkinter import *
import time
import pyautogui


def get_day_dict(day: Day) -> dict:
    return {
        "Entrada": day.interval.start if day.day_type == "REGULAR" else day.day_type,
        "Salida": day.interval.end if day.day_type == "REGULAR" else day.day_type,
    }


def get_cajeros_dataframe(cajeros_list: list[Employee]) -> pd.DataFrame:
    days_of_week_names = ["Lunes", "Martes", "Miércoles",
                          "Jueves", "Viernes", "Sábado", "Domingo"]
    days_of_week_atr_name = [
        "monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

    # Collect dictionaries first
    dict_list = []
    for cajero in cajeros_list:
        cajero_dict = {("Nombre", ""): cajero.name}
        for day_name, day_attr in zip(days_of_week_names, days_of_week_atr_name):
            day_dict = get_day_dict(getattr(cajero.schedule, day_attr))
            for key, value in day_dict.items():
                cajero_dict[(day_name, key)] = value
        dict_list.append(cajero_dict)

    # Create the DataFrame from the list of dictionaries
    cajeros_df = pd.DataFrame(dict_list)

    # Set MultiIndex columns
    cajeros_df.columns = pd.MultiIndex.from_tuples(cajeros_df.columns)

    return cajeros_df

# Función para formatear las columnas de "Entrada" y "Salida"


def format_schedule(df):
    days_of_week_names = ["Lunes", "Martes", "Miércoles",
                          "Jueves", "Viernes", "Sábado", "Domingo"]
    # Aplanar temporalmente el MultiIndex
    df_flat = df.copy()
    df_flat.columns = ['_'.join(col).strip() for col in df_flat.columns.values]

    # Formatear las columnas de "Entrada" y "Salida"
    for day in days_of_week_names:
        for period in ["Entrada", "Salida"]:
            col_name = f"{day}_{period}"
            if col_name in df_flat.columns:
                df_flat[col_name] = df_flat[col_name].apply(
                    lambda x: pd.to_datetime(x).strftime(
                        '%I:%M%p') if not isinstance(x, str) else x
                )

    # Restaurar el MultiIndex
    new_columns = [tuple(col.split('_')) if '_' in col else (col, '')
                   for col in df_flat.columns]
    df_flat.columns = pd.MultiIndex.from_tuples(new_columns)

    return df_flat


def get_day_schedule(df, day):
    output_df = df[["Nombre", day]].xs(day, axis=1, level=0)
    output_df["Nombre"] = df["Nombre"]
    return output_df[["Nombre", "Entrada", "Salida"]]


def format_schedule(df):
    formated_df = df.copy()
    formated_df["Entrada"] = formated_df["Entrada"].apply(
        lambda x: pd.to_datetime(x).strftime(
            '%I:%M%p') if not isinstance(x, str) else x
    )
    formated_df["Salida"] = formated_df["Salida"].apply(
        lambda x: pd.to_datetime(x).strftime(
            '%I:%M%p') if not isinstance(x, str) else x
    )
    return formated_df


class DaySchedule():
    def __init__(self, cajeros_df: pd.DataFrame, day: str):
        self.day = day
        self.cajeros_df = cajeros_df
        self.day_schedule_df = self.get_day_schedule()

    def get_day_schedule(self):
        return get_day_schedule(self.cajeros_df, self.day)

    def format_schedule(self):
        return format_schedule(self.day_schedule_df)

    def get_available_employees(self) -> pd.DataFrame:
        out_df = self.day_schedule_df.copy()
        out_df = out_df[out_df["Entrada"] != "DIA DE DESCANSO"]
        out_df = out_df[out_df["Entrada"] != "VACACIONES"]
        out_df = out_df[out_df["Entrada"] != "PAGO HORAS FERIADO"]

        return out_df


def get_cajeros_df(pdf_path: str, self_amount: int = 6) -> pd.DataFrame:
    employee_list = scrap_pdf(pdf_path, self_amount)
    cajeros_list: list[Employee] = [
        employee for employee in employee_list if employee.category == "CAJERO"]
    return get_cajeros_dataframe(cajeros_list)


def clear():
    os.system('cls' if os.name == 'nt' else 'clear')


def check_path_exists(path):
    return os.path.exists(path)


def get_ean13(cod):
    if len(cod) < 12:
        return None
    if len(cod) > 12:
        cod = cod[:12]
    cod = cod[::-1]
    sum = 0
    for i in range(0, len(cod)):
        if i % 2 == 0:
            sum += int(cod[i])
        else:
            sum += int(cod[i]) * 3
    check_digit = 10 - (sum % 10)
    if check_digit == 10:
        check_digit = 0
    return cod[::-1] + str(check_digit)


def print_menu():
    print('''1. Cargar Archivo (Input ruta)
2. Establecer cantidad de Self
3. Seleccionar Dia (L-D)
4. Mostrar Entradas
5. Mostrar Salidas
6. Entradas-Salidas de la Semana
7. Exportar Entradas y Salidas de la Semana a Excel
8. Entradas-Salidas por día (Clipboard)
9. Tools (Completar EAN)
10. Completar EAN 13 Producto de Peso
11. Exportar a WhatsApp
12. Salir
''')


def print_error():
    clear()
    input('Opción no válida. Presione Enter para continuar...')


def print_header(pdf_path, amount_self, day):
    print(
        f"Menú Principal\t\tArchivo: {pdf_path}\t\t|Self: {amount_self}\t\t|Día: {day}\n")


pdf_path = r"./horarios/Horario 01-07.24.pdf"
amount_self = 6
cajeros_df = get_cajeros_df(pdf_path, amount_self)
day = "Martes"
day_schedule = DaySchedule(cajeros_df, day)
day_list = ["Lunes", "Martes", "Miércoles",
            "Jueves", "Viernes", "Sábado", "Domingo"]

def escribirTxt(file_path, line):
    """Escribe una línea en el archivo especificado."""
    with open(file_path, 'w') as file:
        file.write(line)

def leerTxt(file_path):
    """Lee y devuelve la única línea del archivo especificado."""
    with open(file_path, 'r') as file:
        return file.readline().strip()



def countdown(seconds):
    while seconds > 0:
        time.sleep(1)
        print(seconds)
        seconds -= 1
    print("¡Tiempo terminado!")


def main():

    global pdf_path
    global amount_self
    global cajeros_df
    global day
    global day_schedule
    global day_list

    while True:
        pdf_path = leerTxt("horario.txt")
        clear()
        print_header(pdf_path, amount_self, day)
        print_menu()
        option = input('Ingrese una opción: ')

        if option == '1':
            clear()
            print_header(pdf_path, amount_self, day)
            #Meter gui para abrir archivo
            pdf_path = input(f'Ingrese la ruta del archivo PDF [{pdf_path}]: ')

            #Para arrastrar nomas el archivo
            if pdf_path.startswith('"') and pdf_path.endswith('"'):
                # Eliminar las comillas al inicio y al final
                pdf_path = pdf_path[1:-1]

            #pdf_path = filedialog.askopenfilename()
            if check_path_exists(pdf_path):
                cajeros_df = get_cajeros_df(pdf_path, amount_self)
                input('Archivo cargado correctamente. Presione Enter para continuar...')
                escribirTxt("./horario.txt", pdf_path)
            else:
                print('El archivo no existe')

        elif option == '2':
            clear()
            print_header(pdf_path, amount_self, day)
            amount_self = int(
                input(f'Ingrese la cantidad de Self [{amount_self}]: '))
            cajeros_df = get_cajeros_df(pdf_path, amount_self)
            input(
                'Cantidad de Self actualizada correctamente. Presione Enter para continuar...')

        elif option == '3':
            clear()
            print_header(pdf_path, amount_self, day)
            for i, day in enumerate(day_list):
                print(f"{i+1}. {day}")
            day = day_list[int(input('Ingrese el día (1-7): ')) % 7 - 1]
            day_schedule = DaySchedule(cajeros_df, day)
            input(
                f'Día {day} seleccionado correctamente. Presione Enter para continuar...')

        elif option == '4':
            clear()
            print_header(pdf_path, amount_self, day)
            print(format_schedule(day_schedule.get_available_employees(
            ).sort_values(by="Entrada")[["Nombre", "Entrada", "Salida"]]).to_string(
                index=False
            ))
            input('Presione Enter para continuar...')

        elif option == '5':
            clear()
            print_header(pdf_path, amount_self, day)
            print(format_schedule(day_schedule.get_available_employees(
            ).sort_values(by="Salida")[["Nombre", "Salida", "Entrada"]]).to_string(
                index=False
            ))
            input('Presione Enter para continuar...')

        elif option == '6':
            clear()
            print_header(pdf_path, amount_self, day)
            for _day in day_list:
                print(f"-------------------------------{_day}-------------------------------")
                _day_schedule = DaySchedule(cajeros_df, _day)
                _sorted_schedule = _day_schedule.get_available_employees().sort_values(by="Entrada")[["Nombre", "Entrada", "Salida"]]
                if (_sorted_schedule.empty):
                    print("No hay empleados disponibles")
                else:
                    print(format_schedule(_sorted_schedule).to_string(index=False))
                print("\n\n")
            input('Presione Enter para continuar...')

        elif option == '7':
            clear()
            print_header(pdf_path, amount_self, day)

            for _day in day_list:
                print(f"-------------------------------{_day}-------------------------------")
                _day_schedule = DaySchedule(cajeros_df, _day)
                _available_employees = _day_schedule.get_available_employees()
                
                if _available_employees.empty:
                    print("No hay empleados disponibles")
                else:
                    try:
                        os.makedirs("Exportados")
                    except FileExistsError:
                        pass

                    # Asegurar que las columnas 'Entrada' y 'Salida' sean de tipo datetime
                    _available_employees["Entrada"] = pd.to_datetime(_available_employees["Entrada"])
                    _available_employees["Salida"] = pd.to_datetime(_available_employees["Salida"])

                    # Ordenar por entrada
                    sorted_by_entry = _available_employees.sort_values(by="Entrada")[["Nombre", "Entrada", "Salida"]]
                    sorted_by_entry.columns = ["Nombre_Entrada", "Entrada", "Salida"]

                    # Convertir a texto para evitar "1/01/1900"
                    sorted_by_entry["Entrada"] = sorted_by_entry["Entrada"].dt.strftime('%H:%M')
                    sorted_by_entry["Salida"] = sorted_by_entry["Salida"].dt.strftime('%H:%M')

                    # Ordenar por salida
                    sorted_by_exit = _available_employees.sort_values(by="Salida")[["Nombre", "Salida", "Entrada"]]
                    sorted_by_exit.columns = ["Nombre_Salida", "Salida", "Entrada"]

                    # Convertir a texto para evitar "1/01/1900"
                    sorted_by_exit["Salida"] = sorted_by_exit["Salida"].dt.strftime('%H:%M')
                    sorted_by_exit["Entrada"] = sorted_by_exit["Entrada"].dt.strftime('%H:%M')

                    # Crear un DataFrame combinado con las seis columnas
                    combined_df = pd.concat([sorted_by_entry.reset_index(drop=True), sorted_by_exit.reset_index(drop=True)], axis=1)

                    # Renombrar columnas para el archivo Excel
                    combined_df.columns = ["Nombre", "Entrada", "Salida", "Nombre", "Salida", "Entrada"]

                    # Exportar a un solo archivo Excel
                    combined_df.to_excel(f"./Exportados/{_day}_EntradasSalidas.xlsx", index=False)
                    
                    print("Exportado correctamente")
                print("\n\n")

            input('Presione Enter para continuar...')

        
        elif option == '8':
            clear()
            print_header(pdf_path, amount_self, day)

            event_hours = set()
            for employee in pd.DataFrame(day_schedule.get_available_employees()).iterrows():
                event_hours.add(employee[1]["Entrada"])
                event_hours.add(employee[1]["Salida"])

            event_hours = sorted(list(event_hours))
            if len(event_hours) > 0:
                all_events = []
                print(f"------------------------  {day}  ------------------------")
                all_events.append(f"------------------------  {day}  ------------------------")
                for event_hour in event_hours:
                    entran_list = []
                    salen_list = []
                    print(f"--------------------------{event_hour.strftime('%I:%M%p')}--------------------------")
                    all_events.append(f"--------------------------{event_hour.strftime('%I:%M%p')}--------------------------")
                    for employee in pd.DataFrame(day_schedule.get_available_employees()).iterrows():
                        if employee[1]["Entrada"] == event_hour:
                            entran_list.append(f"\t{employee[1]['Nombre']}")
                    for employee in pd.DataFrame(day_schedule.get_available_employees()).iterrows():
                        if employee[1]["Salida"] == event_hour:
                            salen_list.append(f"\t{employee[1]['Nombre']}")
                    if len(entran_list) > 0:
                        print("Entradas:")
                        all_events.append("Entradas:")
                        for entran in entran_list:
                            print(entran)
                            all_events.append(entran)
                    if len(salen_list) > 0:
                        print("Salidas:")
                        all_events.append("Salidas:")
                        for salen in salen_list:
                            print(salen)
                            all_events.append(salen)
                            
                clipboard.copy("\n".join(all_events))

            else:
                print("No hay empleados disponibles")

            input('Presione Enter para continuar...')

        elif option == '9':
            clear()
            print_header(pdf_path, amount_self, day)
            cod = input("Ingrese el código EAN-12: ")
            ean13 = get_ean13(cod)
            print(f"El código EAN-13 es: {ean13}")
            input("Presione cualquier tecla para continuar...")

        elif option == '10':
            clear()
            print_header(pdf_path, amount_self, day)

            while(True):
                cod_id = input("Ingrese identificador (7 dígitos): ")
                if (len(cod_id) != 7):
                    print("Error: El identificador debe tener 7 dígitos")
                    continue
                try:
                    int(cod_id)
                except:
                    print("Error: El identificador debe ser un número")
                    continue
                break
            
            while(True):
                cod_weight = input("Ingrese el peso (5 dígitos): ")
                if (len(cod_weight) > 5):
                    print("Error: El peso debe tener menos de 5 dígitos")
                    continue
                elif (len(cod_weight) < 5):
                    cod_weight = cod_weight.zfill(5)
                try:
                    int(cod_weight)
                except:
                    print("Error: El peso debe ser un número")
                    continue
                break
            
            ean13 = get_ean13(cod_id + cod_weight)
            print(f"El código EAN-13 es: {ean13}")
            input("Presione cualquier tecla para continuar...")

        elif option == '11':
            clear()
            print("Inicio")
            countdown(5)
            delay = 0.1
            week = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
            for iday in week:
                #clear()
                print_header(pdf_path, amount_self, iday)

                event_hours = set()
                for employee in pd.DataFrame(day_schedule.get_available_employees()).iterrows():
                    event_hours.add(employee[1]["Entrada"])
                    event_hours.add(employee[1]["Salida"])

                event_hours = sorted(list(event_hours))
                if len(event_hours) > 0:
                    all_events = []
                    print(f"------------------------  {iday}  ------------------------")
                    all_events.append(f"------------------------  {iday}  ------------------------")
                    for event_hour in event_hours:
                        entran_list = []
                        salen_list = []
                        print(f"--------------------------{event_hour.strftime('%I:%M%p')}--------------------------")
                        all_events.append(f"--------------------------{event_hour.strftime('%I:%M%p')}--------------------------")
                        for employee in pd.DataFrame(day_schedule.get_available_employees()).iterrows():
                            if employee[1]["Entrada"] == event_hour:
                                entran_list.append(f"\t{employee[1]['Nombre']}")
                        for employee in pd.DataFrame(day_schedule.get_available_employees()).iterrows():
                            if employee[1]["Salida"] == event_hour:
                                salen_list.append(f"\t{employee[1]['Nombre']}")
                        if len(entran_list) > 0:
                            print("Entradas:")
                            all_events.append("Entradas:")
                            for entran in entran_list:
                                print(entran)
                                all_events.append(entran)
                        if len(salen_list) > 0:
                            print("Salidas:")
                            all_events.append("Salidas:")
                            for salen in salen_list:
                                print(salen)
                                all_events.append(salen)
                                
                    clipboard.copy("\n".join(all_events))
                    countdown(delay)
                    pyautogui.hotkey('ctrl', 'v')
                    countdown(delay)
                    pyautogui.press('enter')
                    countdown(delay)
            input("Mensajes Enviados")

        elif option == '12':
            break            

        else:
            pass
            input('Presione Enter para continuar...')


if __name__ == "__main__":
    main()

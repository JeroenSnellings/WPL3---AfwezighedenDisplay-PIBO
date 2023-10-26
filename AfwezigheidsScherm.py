import os
import time
import tkinter as tk
import threading
import openpyxl
from screeninfo import get_monitors

# Variabelen
monitor_window = None

# Resolutie van de monitor(s) bepalen
monitors = get_monitors()
if len(monitors) > 1:
    current_monitor = monitors[1]  # Index 1 = monitor 2, index 0 = monitor 1 (main display)
    screen_width = current_monitor.width
    screen_height = current_monitor.height
else:
    current_monitor = monitors[0]
    screen_width = monitors[0].width
    screen_height = monitors[0].height


# FUNCTIE --> Leest de inhoud van het Excel bestand
def read_excel_file():
    try:
        wb = openpyxl.load_workbook('Afwezigheden.xlsx')
        sheet = wb.active
        data = []
        data.append("\n")
        for row in sheet.iter_rows(values_only=True):
            if row is not None:
                rowtext = ""
                for cell in row:
                    if cell is not None and str(cell).upper() != "AFWEZIGHEDEN LEERKRACHTEN":
                        rowtext += str(cell + "\n")
                data.append(str(rowtext))
        wb.close()
        return data
    except Exception as e:
        return [["Error: kan data van Excel bestand niet lezen!"]]


# FUNCTIE --> Update het scherm met de nieuwe informatie
def update_window():
    global monitor_window
    if monitor_window:
        # verwijder het oude venster en maak een nieuwe aan
        monitor_window.destroy()
        monitor_window = tk.Toplevel()
        monitor_window.title("Leerkracht Afwezigheidsmonitor")
       # monitor_window.geometry(f"{screen_width}x{screen_height}+{current_monitor.x}+{current_monitor.y}")
        monitor_window.configure(bg="black")
        monitor_window.attributes('-fullscreen', True)
        monitor_window.attributes('-topmost', True)
        monitor_window.bind('<Escape>', exit_fullscreen)  # Bind Esc key om uit fullscreen te gaan

    data = read_excel_file()

    header_frame = tk.Frame(monitor_window, bg="black")
    header_frame.pack()

    # Header label voor de titel
    lector_Naam_Label = tk.Label(
        header_frame,
        text="Afwezigheden Leerkrachten PIBO",
        font=("Arial", 46, "bold"),
        bg="black",
        fg="white",
        anchor="e"
    )
    lector_Naam_Label.pack(side="top")

    data_frame = tk.Frame(monitor_window, bg="black")
    data_frame.pack()

    # label aanmaken en weergeven op de grid
    data_label = tk.Label(data_frame, text="".join(["".join(map((str), row)) for row in data]),
                          font=("Arial", 40, "bold"), bg="black", fg="white", anchor="e")
    data_label.grid(row=1, column=1)


# FUNCTIE --> Nodig om uit de fullscreen te ontsnappen voor debugging
def exit_fullscreen(event):
    global monitor_window
    monitor_window.attributes('-fullscreen', False)
    monitor_window.attributes('-topmost', False)


# FUNCTIE --> Start een thread die het excel bestand in de gaten houdt
def check_excel_file():
    last_modified_time = os.path.getmtime('Afwezigheden.xlsx')

    while True:
        time.sleep(1)
        if os.path.getmtime('Afwezigheden.xlsx') > last_modified_time:
            update_window()
            last_modified_time = os.path.getmtime('Afwezigheden.xlsx')


# Aanmaken van tkinter window
monitor_window = tk.Toplevel()
monitor_window.title("Excel Viewer")
monitor_window.configure(bg="black")

# thread aanmaken die het excel bestand in de gaten houdt

file_check_thread = threading.Thread(target=check_excel_file)
file_check_thread.daemon = True
file_check_thread.start()

# Initialiseren van de window
update_window()

# Start de Tkinter main loop
monitor_window.mainloop()

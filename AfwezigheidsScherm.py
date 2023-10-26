# *******************************************************************
# *                                                                 *
# *                AfwezigheidsScherm                               *
# *                                                                 *
# *  Auteur: Jeroen Snellings                                       *
# *  Datum: 25 oktober 2023                                         *
# *  Versie: 1.1                                                    *
# *                                                                 *
# *  Beschrijving:                                                  *
# *  Dit programma is verantwoordelijk voor het vertonen van de     *
# *  afweigheden van de docenten van de PIBO op het scherm in de    *
# *  inkomhal.                                                      *
# *                                                                 *
# *  Auteursrecht (c) 2023 Jeroen Snellings                         *
# *                                                                 *
# *  Hierbij wordt toestemming verleend, kosteloos, aan eenieder    *
# *  die een kopie van deze software en de bijbehorende             *
# *  documentatiebestanden (de "Software") verkrijgt, om de         *
# *  Software zonder beperkingen te gebruiken, inclusief, maar      *
# *  niet beperkt tot, het recht om kopieën van de Software te      *
# *  gebruiken, te kopiëren, te wijzigen, samen te voegen, te       *
# *  publiceren, te distribueren, onder de volgende voorwaarden:    *
# *                                                                 *
# *  De bovenstaande auteursrechtverklaring en deze                 *
# *  toestemmingsverklaring moeten in alle kopieën of               *
# *  aanzienlijke delen van de Software worden opgenomen.           *
# *                                                                 *
# *  DE SOFTWARE WORDT GELEVERD "AS IS", ZONDER ENIGE               *
# *  GARANTIE, EXPLICIET OF IMPLICIET, INCLUSIEF, MAAR NIET         *
# *  BEPERKT TOT, DE GARANTIES VAN VERKOOPBAARHEID, GESCHIKTHEID    *
# *  VOOR EEN BEPAALD DOEL EN NIET-INBREUK. IN GEEN GEVAL ZULLEN    *
# *  DE AUTEURS OF HOUDERS VAN HET AUTEURSRECHT AANSPRAKELIJK       *
# *  ZIJN VOOR ENIGE CLAIM, SCHADE OF ANDERE AANSPRAKELIJKHEID,     *
# *  ONGEACHT IN EEN ACTIE VAN CONTRACT, ONRECHT OF ANDERSZINS,     *
# *  VOORTVLOEIEND UIT, OF IN VERBAND MET DE SOFTWARE OF HET        *
# *  GEBRUIK OF ANDERE HANDELINGEN IN DE SOFTWARE.                  *
# *                                                                 *
# *                                                                 *
# *  VERKOPEN VAN DEZE SOFTWARE IS NIET TOEGESTAAN                  *
# *                                                                 *
# *******************************************************************

import os
import time
import tkinter as tk
import threading
import openpyxl
from screeninfo import get_monitors

# Variabelen
monitor_window = None

# Determine the screen resolution of the second monitor
monitors = get_monitors()
if len(monitors) > 1:
    second_monitor = monitors[1]  # Use index 1 to get the second monitor
    screen_width = second_monitor.width
    screen_height = second_monitor.height
else:
    screen_width = monitor_window.winfo_screenwidth()
    screen_height = monitor_window.winfo_screenheight()



# FUNCTIE --> Leest de inhoud van het Excel bestand
def read_excel_file():
    try:
        wb = openpyxl.load_workbook('Afwezigheden.xlsx')
        sheet = wb.active
        data = []
        for row in sheet.iter_rows(values_only=True):
            if row is not None:
                rowtext = ""
                for cell in row:
                    if cell is not None:
                        if str(cell).upper() == "AFWEZIG":
                            rowtext += str("Afwezig")
                        else:
                            rowtext += str(cell + " --> ")
                data.append(str(rowtext))
                # lege_Rij = [str(cell + " |  | ") if cell is not None else "" for cell in row]
                # # kijkt of de rij niet leeg is
                # if any(lege_Rij):
                #     data.append(lege_Rij)
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
        monitor_window.title("Lector Afwezigheidsmonitor")
        monitor_window.geometry(f"{screen_width}x{screen_height}+{second_monitor.x}+{second_monitor.y}")
        monitor_window.configure(bg="black")
    #    monitor_window.attributes('-fullscreen', True)
        monitor_window.attributes('-topmost', True)
        monitor_window.bind('<Escape>', exit_fullscreen)  # Bind Esc key om uit fullscreen te gaan

    data = read_excel_file()

    header_frame = tk.Frame(monitor_window, bg="black")
    header_frame.pack()

    # Header label voor de titel
    lector_Naam_Label = tk.Label(
        header_frame,
        text="Afwezigheden Lectoren PIBO",
        font=("Arial", 46, "bold"),
        bg="black",
        fg="white",
        anchor="center"
    )
    lector_Naam_Label.pack(side="top")


    data_frame = tk.Frame(monitor_window, bg="black")
    data_frame.pack()

    # Create a label to display the data using 'grid'
    data_label = tk.Label(data_frame, text="\n".join(["".join(map((str), row)) for row in data]),
                          font=("Arial", 30, "bold"), bg="black", fg="white", anchor="center")
    data_label.grid(row=1, column=1)

# FUNCTIE --> Nodig om uit de fullscreen te ontsnappen voor debugging
def exit_fullscreen(event):
    global monitor_window
    monitor_window.attributes('-fullscreen', False)
    monitor_window.attributes('-topmost', False)

# FUNCTIE --> Start een thread die het excel bestand in de gaten houdt
def check_excel_file():
    # Store the initial timestamp of the Excel file
    last_modified_time = os.path.getmtime('Afwezigheden.xlsx')

    while True:
        # Check if the file has been modified

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

"""Genera ejemplo_sucio.xlsx con datos INVENTADOS y desordenados a propósito.

Sirve para probar y enseñar el limpiador sin usar datos reales de nadie.
"""
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font

CABECERAS = [" Nombre ", "EMAIL", "Teléfono", "Fecha alta", "Importe", "ciudad"]

FILAS = [
    ["  ana garcía lópez ", "Ana.Garcia@EjemploMail.com ", "600-123-456", "05/03/2024", "1.234,50 €", "madrid"],
    ["LUIS MARTÍN", "luis.martin@ejemplomail.com", "+34 611 222 333", "2024-03-07", "89 EUR", "  Sevilla "],
    [None, None, None, None, None, None],  # fila vacía
    ["Marta Ruiz", "marta.ruiz@ejemplomail.com", "622.333.444", datetime(2024, 3, 9), 450.0, "MADRID"],
    ["ana garcía lópez", "ana.garcia@ejemplomail.com", "600123456", "5-3-2024", "1234,5", "Madrid"],  # duplicado de Ana
    ["Pedro  Gómez", "pedro.gomez@ejemplomail", "633 444 555", "12/03/2024", "99.99", "valencia"],  # email inválido
    ["Lucía Fernández", "lucia.f@ejemplomail.com", "644555666", "31/02/2024", "300", "Bilbao"],  # fecha imposible
    ["jorge  sanz", "JORGE.SANZ@EJEMPLOMAIL.COM", "655 666 777", 45365, "2.500", "barcelona"],  # fecha como nº de Excel
    ["Elena Torres", "elena.torres@ejemplomail.com", "666-777-888", "2024/03/15", "abc", "Zaragoza"],  # importe inválido
    ["LUIS MARTÍN", "LUIS.MARTIN@ejemplomail.com", "611222333", "07/03/2024", "89,00", "sevilla"],  # duplicado de Luis
    ["Carmen Vidal", "carmen.vidal@ejemplomail.com", "677 888 99", "20/03/2024", "1 200,00 €", "Málaga"],  # teléfono corto
]


def main(ruta="ejemplo_sucio.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Clientes"
    ws.append(CABECERAS)
    for fila in FILAS:
        ws.append(fila)
    for celda in ws[1]:
        celda.font = Font(name="Arial", bold=True)
    for fila in ws.iter_rows(min_row=2):
        for celda in fila:
            celda.font = Font(name="Arial")
    for col, ancho in zip("ABCDEF", (24, 32, 18, 14, 14, 14)):
        ws.column_dimensions[col].width = ancho
    wb.save(ruta)
    print(f"Creado {ruta} con {len(FILAS)} filas de ejemplo")


if __name__ == "__main__":
    main()

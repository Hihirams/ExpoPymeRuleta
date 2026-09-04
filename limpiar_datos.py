# -*- coding: utf-8 -*-
"""Limpia por completo la base de datos local (respuestas.db):

- Borra TODAS las respuestas (encuestas contestadas).
- Borra todos los premios entregados y el contador por hora (prize_events).
- Restablece la configuracion (premios de la ruleta, PIN, "Premios por hora",
  etc.) a los valores por defecto del codigo.

Uso (doble clic en LIMPIAR_DATOS.bat o desde la terminal):
    python limpiar_datos.py

NOTA: solo limpia el SERVIDOR (respuestas.db) de la laptop. Lo que haya
quedado guardado en los CELULARES (localStorage/IndexedDB) se limpia por
dispositivo: engrane (rueda) -> "Borrar historico", o abriendo la pagina en
una ventana de incognito.
"""

import os
import sqlite3

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE_DIR, "respuestas.db")

# Siempre en este orden: primero datos, luego configuracion.
TABLAS = ("respuestas", "prize_events", "app_config")


def main():
    if not os.path.exists(DB):
        print("No existe respuestas.db todavia (nada que limpiar).")
        return

    try:
        con = sqlite3.connect(DB)
    except sqlite3.Error as ex:
        print("No se pudo abrir la base de datos: %s" % ex)
        return

    borradas = {}
    try:
        for t in TABLAS:
            try:
                cur = con.execute("DELETE FROM %s" % t)
                borradas[t] = cur.rowcount
            except sqlite3.Error:
                borradas[t] = 0
        con.commit()
    except sqlite3.Error as ex:
        print("Ocurrio un error al limpiar: %s" % ex)
    finally:
        # Compactar el archivo es opcional; si el servidor esta corriendo y
        # bloquea el archivo, se omite silenciosamente (ya se borro todo).
        try:
            con.execute("VACUUM")
            con.commit()
        except sqlite3.Error:
            pass
        con.close()

    print("Base de datos: %s" % DB)
    print("  - respuestas     (encuestas): %d borrada(s)" % borradas.get("respuestas", 0))
    print("  - prize_events   (premios/hora): %d borrada(s)" % borradas.get("prize_events", 0))
    print("  - app_config     (config): %d borrada(s)" % borradas.get("app_config", 0))
    print()
    print("Listo: encuestas, premios entregados y cupo por hora quedaron en 0.")
    print("El servidor reconstruye la base por si sola al arrancar.")


if __name__ == "__main__":
    main()
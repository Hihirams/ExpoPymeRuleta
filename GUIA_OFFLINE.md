# Guía: encuesta + ruleta con QR, **100% sin internet**

Esta guía explica cómo poner un **QR** en el evento para que las personas lo
escaneen con **su propio celular**, llenen la encuesta y la ruleta, y que cada
respuesta llegue a **una sola base de datos** en tu laptop — **sin internet en
ningún lado** (ni la laptop ni los celulares necesitan datos).

## ¿Cómo funciona? (la idea en 30 segundos)

Internet y "red WiFi" son cosas distintas. Para que los celulares hablen con la
laptop solo hace falta que estén en **la misma red local**. Esa red la crea la
propia laptop (su *Mobile Hotspot*), aislada del mundo.

```
   Sin señal celular, sin cable, sin nada externo
   ┌────────────────────────────────────────────────┐
   │   Burbuja local (la crea tu laptop)             │
   │                                                 │
   │   [Laptop]  ←── WiFi local ──→  [Celulares]     │
   │   192.168.137.1                 escanean el QR  │
   │   sirve la encuesta             llenan y envían │
   │   guarda TODO en respuestas.db                  │
   └────────────────────────────────────────────────┘
```

- El **QR** apunta a la IP local de la laptop, p. ej. `http://192.168.137.1:8080`.
- Cada respuesta viaja del celular a la laptop y se guarda en `respuestas.db`.
- Si el WiFi falla un momento, el celular **guarda la respuesta y la reintenta
  solo** (no se pierde nada).

---

## 1) Preparación (UNA sola vez, con internet, en tu oficina)

Instala las dependencias mientras **todavía tienes internet**:

```bash
pip install -r requirements.txt
```

> `http.server` y `sqlite3` ya vienen con Python. Solo se instalan `openpyxl`
> (para el Excel) y `qrcode` (para dibujar el QR). Después de esto, ya **no
> necesitas internet nunca más**.

Prueba que todo funciona (opcional pero recomendado):

```bash
python test_servidor.py
```

Debe decir `OK` al final.

---

## 2) Crear la red WiFi local (sin internet)

Windows **no deja encender** el *Mobile Hotspot* si la laptop no tiene ninguna
conexión ("We can't set up mobile hotspot because your PC doesn't have an
Ethernet, Wi-Fi, or cellular data connection"). Para solucionarlo hay un
adaptador **loopback** que le da a Windows una "conexión" para compartir.

### 2.1 Preparar el adaptador (ya hecho, una sola vez)

Doble clic en **`CONFIGURAR_HOTSPOT.bat`** → di **"Sí"** al aviso de permisos.
Instala y configura solo el adaptador **HotspotLoopback**. (Ya quedó instalado.)

### 2.2 Encender el Mobile Hotspot

1. Abre **Configuración → Red e Internet → Mobile hotspot**.
2. **"Compartir mi conexión de Internet desde"** → elige **HotspotLoopback**.
3. **"Compartir a través de"** → **Wi-Fi**.
4. **Enciende** el interruptor (Off → On).
5. Anota el **Nombre de red** y la **Contraseña** (los celulares la usarán).

> Con el *Mobile Hotspot* de Windows, la laptop **siempre** tiene la IP
> `192.168.137.1`. Eso hace que la URL y el QR sean **estables**: puedes
> imprimir el QR una vez y reutilizarlo en cada evento.

> ⚠️ **Si el interruptor sigue sin encender** (algunas versiones muy nuevas de
> Windows 11 exigen un perfil "con internet" real): no pierdas tiempo, usa un
> **celular Android como hotspot** (modo avión + WiFi + hotspot, sin datos) o un
> **router WiFi**. La laptop se conecta a esa red, corres el servidor y el panel
> te da el QR con la IP correcta automáticamente. Todo lo demás es idéntico.

> **Para quitar el adaptador** después: Administrador de dispositivos →
> Adaptadores de red → clic derecho en *Microsoft KM-TEST Loopback Adapter* →
> **Desinstalar dispositivo**.

---

## 3) Iniciar el servidor

Haz **doble clic** en **`INICIAR_SERVIDOR.bat`**.

- Se abre una ventana negra con la URL y un **QR** dibujado en texto.
- Se abre solo el **panel de control** en el navegador:
  `http://localhost:8080/panel`

O desde la terminal:

```bash
python servidor.py --port 8080
```

Deja esa ventana **abierta** durante todo el evento y la laptop **enchufada**.

---

## 4) El QR para que la gente escanee

Tienes tres formas de obtener el QR (todas offline):

1. **Panel** (`http://localhost:8080/panel`): muestra el QR grande y en vivo.
   Proyéctalo o ponlo en una pantalla.
2. **Imagen** `http://localhost:8080/qr.png`: ábrela, guárdala e **imprímela**
   para pegarla en el stand.
3. **Terminal**: el QR también se dibuja en la ventana negra al iniciar.

> **Recomendado:** imprime el QR en papel y pégalo. Como la IP del hotspot
> siempre es `192.168.137.1`, el mismo papel sirve para todos los eventos.

---

## 5) Qué hacen las personas

1. Encienden **WiFi** en su celular y se conectan a la red de tu laptop
   (la del hotspot). **No necesitan datos ni plan.**
2. Escanean el QR (o escriben la URL en el navegador).
3. Llenan la encuesta, giran la ruleta y presionan **Guardar y terminar**.
4. Listo: la respuesta ya está en tu laptop.

> ⚠️ **Aviso "Red sin internet":** como la red no tiene salida a internet,
> algunos celulares muestran ese aviso o preguntan si quieren seguir conectados.
> Hay que decirles **"Sí, mantener conexión"**. La página abre igual. Si un
> celular insiste en desconectarse, que desactive los "datos móviles" un momento.

---

## 6) Ver resultados y exportar el Excel

En la laptop, en el **panel** (`http://localhost:8080/panel`):

- Ves el **conteo en vivo** de respuestas y premios (se actualiza cada 4 s).
- Botón **"Exportar Excel"** → descarga el archivo con las **49 columnas** de
  siempre (mismo formato que *"Encuesta de satisfacción Expo PyME.xlsx"*), listo
  para tu script `descargar_encuestas.py` / procesamiento habitual.

También puedes exportar entrando a `http://localhost:8080/export.xlsx`.

---

## 6b) Controlar la probabilidad de premio EN VIVO (todos los celulares)

La configuración de la ruleta vive en el **servidor**, así que lo que cambies en
el admin de la laptop se aplica a **todos los celulares** en ≤15 segundos.

1. En la laptop (página servida), toca el **engrane** → escribe el **PIN**.
2. En la tabla de premios, cambia la columna **Peso**:
   - Peso más alto en los segmentos de premio = **más gente gana**.
   - Peso más bajo = **menos gente gana**.
   - `probabilidad = (suma de pesos de los segmentos "Premio") ÷ (suma de todos los pesos)`.
3. El cambio se guarda solo y **cada celular lo toma en su siguiente giro**.

Ejemplo: si sientes que están ganando demasiado, baja el peso de los premios
(p. ej. de 15/8 a 5/3) y en segundos todos giran con menos probabilidad. Está
protegido con tu **PIN**: solo quien lo sepa puede cambiarlo.

> Nota: el **stock** (cantidad) se cuenta por dispositivo; para el control de
> "cuánta gente gana" usa el **Peso**, que sí es central.

## 7) Dónde están los datos y respaldos

- **Todo** se guarda de forma centralizada en **`respuestas.db`** (SQLite), en la
  misma carpeta del servidor, en la laptop.
- Cada respuesta se guarda apenas llega; no depende de ningún celular.
- **Respaldo:** exporta el Excel de vez en cuando (cada hora, por ejemplo) y
  copia `respuestas.db` a una USB. Así, aunque algo pase con la laptop, tienes
  copias.
- Cuando vuelvas a tener internet, si quieres, sube el Excel a la nube.

---

## 8) Solución de problemas

| Problema | Solución |
|---|---|
| El celular no abre la página | Verifica que esté conectado a la **WiFi del hotspot** (no a otra red). Prueba escribir la URL a mano. |
| Sale "Red sin internet" | Es normal (no hay internet). Elegir **"Mantener conexión"**. |
| El QR no funciona | Usa otra IP: el servidor lista todas las IP de la laptop al iniciar. En el evento debe ser `192.168.137.1`. |
| "openpyxl no instalado" | Corre `pip install -r requirements.txt` (con internet, una vez). |
| No se ve el QR en la terminal | No pasa nada: usa el panel `/panel` o `/qr.png`. |
| Muchos celulares a la vez se traban | El *Mobile Hotspot* de Windows aguanta ~8 equipos. Para 20–50 celulares, usa un **router de viaje** barato conectado a la laptop por cable, y en el QR pon la IP que ese router le dé a la laptop. |

---

## Archivos del proyecto

| Archivo | Qué es |
|---|---|
| `index.html` | La encuesta + ruleta (responsive: celular, tablet, PC). Ahora **envía** cada respuesta al servidor. |
| `servidor.py` | El servidor offline: sirve la página, recibe respuestas, guarda en SQLite y exporta a Excel. |
| `INICIAR_SERVIDOR.bat` | Doble clic para iniciar todo en Windows. |
| `test_servidor.py` | Pruebas automáticas del servidor. |
| `requirements.txt` | Dependencias (`openpyxl`, `qrcode`). |
| `respuestas.db` | Base de datos con todas las respuestas (se crea sola). |

> **Modo kiosco antiguo:** si abres `index.html` como archivo (doble clic, sin
> servidor), sigue funcionando como antes — guarda solo en ese dispositivo. El
> envío al servidor **solo** se activa cuando la página se abre desde
> `http://…` (es decir, servida por `servidor.py`).

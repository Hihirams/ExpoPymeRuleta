# Guía: encuesta + ruleta con QR, **100% sin internet**

Esta guía explica cómo poner un **QR** en el evento para que las personas lo
escaneen con **su propio celular**, llenen la encuesta y la ruleta, y que cada
respuesta llegue a **una sola base de datos** en tu laptop — **sin internet en
ningún lado** (ni la laptop ni los celulares necesitan datos).

## ¿Cómo funciona? (la idea en 30 segundos)

Internet y "red WiFi" son cosas distintas. Para que los celulares hablen con la
laptop solo hace falta que estén en **la misma red local**. Esa red la crea un
**celular en modo hotspot** (o un router), aislada del mundo; la **laptop se
conecta a esa red** y corre el servidor.

```
   Sin señal celular real, sin cable, sin nada externo
   ┌─────────────────────────────────────────────────────────┐
   │   Burbuja local (la crea un CELULAR en hotspot o router) │
   │                                                          │
   │   [Celular hotspot]  ← WiFi →  [Laptop]  ← WiFi → [Cels] │
   │    crea la red                  servidor         escanean│
   │                                 respuestas.db    y envían│
   └─────────────────────────────────────────────────────────┘
```

- El **QR** apunta a la IP que tomó la laptop en esa red (el servidor la detecta
  y la pone sola en el panel).
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

⚠️ **La laptop NO crea el hotspot.** En este equipo, el *Mobile Hotspot* de
Windows no enciende sin internet. Así que la red WiFi la crea **otro aparato**
(un celular o un router) y **la laptop se conecta a esa red como un cliente
más**. Ese es el orden correcto.

### Opción recomendada: un celular como hotspot (gratis, sin datos)

1. En un celular **Android**: activa **modo avión**, luego enciende **WiFi** y
   el **Punto de acceso / Hotspot**. No necesita SIM ni datos. (En iPhone
   funciona si tiene SIM.) → **ese celular crea la red.**
2. Conecta **la laptop** a ese hotspot. *(La laptop es cliente; NO enciende
   ningún hotspot propio.)*
3. Conecta los **celulares de los participantes** al **mismo** hotspot.
4. Anota el **Nombre de red (SSID)** y la **Contraseña** del hotspot: los usarás
   para el 2º QR (ver sección 4).

### Alternativa para mucha gente: un router WiFi

Cualquier router (o *travel router*) crea la red sin internet y aguanta 32–64+
dispositivos. La laptop se conecta al router igual que arriba. Es lo más estable
si esperas mucha gente.

> La laptop tomará una IP como `192.168.x.x` (o `172.20.10.x` si el hotspot es
> de iPhone). **El servidor la detecta solo** y el panel arma el QR con esa IP;
> no tienes que configurar nada.

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

**Para mostrar también un QR que conecte al WiFi** (recomendado), arranca con el
nombre y clave de la red (la del celular/router de la sección 2):

```bash
python servidor.py --ssid "NombreDeLaRed" --wifi-pass "laClave"
```

Así el panel muestra **2 QR**: uno para unirse al WiFi y otro para abrir la
encuesta.

Deja esa ventana **abierta** durante todo el evento y la laptop **enchufada**.

---

## 4) Los QR para que la gente escanee

Hay **dos** QR (ambos en el panel, todo offline):

1. **QR de WiFi** (si arrancaste con `--ssid`): lo escanean y su celular **se une
   a la red** sin teclear la clave.
2. **QR de la encuesta**: lo escanean y **abre la encuesta**.

Dónde verlos:
- **Panel** (`http://localhost:8080/panel`): muestra los QR grandes y el conteo
  en vivo. Proyéctalo o ponlo en una pantalla.
- **Imagen** `http://localhost:8080/qr.png` (encuesta) y `/wifi-qr.png` (WiFi):
  ábrelas e **imprímelas** para pegarlas en el stand.

> **Nota:** la IP de la laptop (y por tanto el QR de la encuesta) **cambia** según
> la red a la que te conectes. Por eso lo mejor es **mostrar el panel en vivo** en
> una pantalla; si prefieres imprimir, regenera el QR cuando cambie la red.

---

## 5) Qué hacen las personas

1. Encienden **WiFi** y se conectan al **mismo hotspot** (el del celular/router
   de la sección 2) — escaneando el **QR de WiFi** o eligiéndolo a mano.
   **No necesitan datos ni plan.**
2. Escanean el **QR de la encuesta** (o escriben la URL en el navegador).
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
| El QR no funciona | El servidor lista **todas las IP** de la laptop al iniciar; usa la que empiece por la red del hotspot (`192.168.x.x` o `172.20.10.x`). El panel arma el QR con la correcta. |
| "openpyxl no instalado" | Corre `pip install -r requirements.txt` (con internet, una vez). |
| No se ve el QR en la terminal | No pasa nada: usa el panel `/panel` o `/qr.png`. |
| Muchos celulares a la vez se traban | El hotspot de un celular aguanta ~8–15 equipos. Para 20–50 celulares, usa un **router WiFi** (aguanta 32–64+). La laptop se conecta a él igual. |

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

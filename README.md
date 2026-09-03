# Expo PyME · Encuesta + Ruleta (offline)

Encuesta de satisfacción con ruleta de premios que funciona **100% sin
internet**. Los participantes escanean un QR con su propio celular, llenan la
encuesta, giran la ruleta y cada respuesta se guarda **centralizada** en una
laptop. Al final se exporta a Excel (49 columnas, formato SurveyMonkey).

## Puesta en marcha (en cualquier laptop)

```bash
git clone https://github.com/Hihirams/ExpoPymeRuleta.git
cd ExpoPymeRuleta
pip install -r requirements.txt        # una sola vez, con internet
```

Necesitas **Python 3.8+**. `http.server` y `sqlite3` ya vienen con Python; solo
se instalan `openpyxl` (Excel) y `qrcode` (QR).

## Correrlo

- **Windows:** doble clic en `INICIAR_SERVIDOR.bat`.
- **Cualquier SO / terminal:**
  ```bash
  python servidor.py --port 8080
  ```
  Para mostrar además un QR que conecta al WiFi:
  ```bash
  python servidor.py --ssid "NombreDeLaRed" --wifi-pass "laClave"
  ```

Luego abre el **panel** en la laptop: `http://localhost:8080/panel` (QR, conteo
en vivo y botón para exportar el Excel).

## La red WiFi (sin internet)

La laptop **no** crea el hotspot: la red la crea **un celular en modo hotspot**
(Android en modo avión, sin datos) o un **router**, y la **laptop se conecta a
esa red** como cliente. Los participantes se unen al mismo WiFi. El servidor
detecta la IP solo y arma el QR en el panel.

👉 Guía completa paso a paso: **[GUIA_OFFLINE.md](GUIA_OFFLINE.md)**.

## Cómo funciona

- `servidor.py` sirve `index.html`, recibe cada respuesta por `POST /api/respuesta`
  y la guarda en `respuestas.db` (SQLite). Exporta en `GET /export.xlsx`.
- La ruleta y sus probabilidades se controlan **en vivo** desde el admin de la
  laptop (engrane → PIN) y se aplican a **todos** los celulares (config central
  en el servidor). PIN por defecto: `1234` (cámbialo).
- Offline-first: si el WiFi falla, el celular reintenta el envío solo.
- Sin servidor (p. ej. abriendo el HTML suelto) funciona como kiosco: guarda en
  el propio dispositivo.

## Pruebas

```bash
python test_servidor.py
```

## Datos

`respuestas.db` (la base con las respuestas) **no** se sube al repo; se crea sola
al correr el servidor. Respáldala copiándola a una USB durante el evento.

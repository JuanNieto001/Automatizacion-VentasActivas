# Verificación de ventas en AC

Automatiza el paso a paso de `PASO A PASO AC.docx`: toma los números de la
columna **E** del Excel de ventas, los consulta en **AC Administración de
Clientes**, y arma un reporte con cuáles quedaron **activas**, cuáles **no**,
y **por qué** las que no lo están.

---

## Uso rápido

1. Copia el Excel de ventas en la carpeta `entrada\`.
2. Doble clic en **`Ejecutar.bat`**.
3. **No toques el mouse ni el teclado** mientras corre (AC se maneja con clics
   reales sobre la pantalla).
4. Al terminar, abre el CSV que queda en `salida\`.

La primera vez pedirá la clave de red. Para no escribirla cada vez:

```powershell
.\Automatizar-Ventas.ps1 -GuardarClave
```

Queda cifrada en `config\clave.dat` con tu usuario de Windows: **solo se puede
descifrar con tu cuenta en este mismo equipo**, y nunca se guarda en texto plano.

---

## Otras formas de ejecutarlo

```powershell
# Un número suelto (prueba rápida)
.\Automatizar-Ventas.ps1 -Numeros 3001234567

# Varios números sin Excel
.\Automatizar-Ventas.ps1 -Numeros 3001234567,3009876543

# Un archivo específico
.\Automatizar-Ventas.ps1 -Archivo "C:\ruta\ventas.xlsx"

# Otra columna u otra fila de inicio
.\Automatizar-Ventas.ps1 -Archivo ".\entrada\ventas.xlsx" -Columna F -FilaInicio 3

# Traer el historial de TODAS las líneas, no solo de las no activas (más lento)
.\Automatizar-Ventas.ps1 -Archivo ".\entrada\ventas.xlsx" -HistorialSiempre

# Varias instancias de AC en paralelo (más rápido para listas grandes)
.\Automatizar-Ventas.ps1 -Archivo ".\entrada\ventas.xlsx" -Instancias 4
```

| Parámetro | Para qué sirve | Por defecto |
|---|---|---|
| `-Archivo` | Excel/CSV/TXT a procesar | el más reciente de `entrada\` |
| `-Numeros` | Números sueltos, sin archivo | — |
| `-Columna` | Columna del Excel con los números | `E` |
| `-FilaInicio` | Primera fila con datos (salta el encabezado) | `2` |
| `-Hoja` | Nombre de la hoja del libro | la primera |
| `-HistorialSiempre` | Consulta el historial de todas las líneas | solo las no activas |
| `-Instancias` | Cuántas copias de AC consultan en paralelo (1-12) | `1` |
| `-BaseDatos` | Base de AC | `AC_PRODUCCION` |
| `-GuardarClave` | Guarda la clave cifrada y sale | — |
| `-Salida` | Ruta del CSV de salida | `salida\ventas_<fecha>.csv` |

Acepta `.xlsx`, `.xlsm`, `.csv` y `.txt`. **No necesita Excel instalado**
(este equipo no lo tiene): el `.xlsx` se lee directamente.

Los números se normalizan solos: quita espacios y guiones, y el indicativo
`57` si viene. Descarta lo que no queden 10 dígitos y elimina repetidos.

---

## Qué trae el reporte

| Columna | Contenido |
|---|---|
| `NUMERO` | Número consultado |
| `ENCONTRADO` | `SI` / `NO` / `ERROR` |
| `ACTIVA` | `SI` cuando el ESTADO de la grilla es `Activo` |
| `ESTADO` | Estado tal cual lo muestra AC |
| `MOTIVO` | Motivo del último movimiento del historial |
| `FECHA_ESTADO` | Fecha "válido desde" de ese movimiento |
| `USUARIO_ESTADO` | Usuario que lo generó |
| `NOMBRE`, `CUSTCODE`, `PLAN`, `TECNOLOGIA`, `TIPO_CLIENTE`, `CENTRO_COSTOS` | Datos de la línea |
| `HISTORIAL_COMPLETO` | Todos los movimientos, separados por `\|\|` |
| `OBSERVACION` | Avisos: no encontrado, varias líneas, errores, discrepancias |
| `CONSULTADO` | Fecha y hora de la consulta |

Además, en `salida\capturas_<fecha>\` queda una captura de la pantalla
HISTORIAL de cada línea revisada, como evidencia.

---

## Cómo trabaja (y por qué así)

Va en dos pasadas:

- **Pasada 1 — estado de cada número.** Busca por criterio MIN/MSISDN,
  selecciona la línea y lee la grilla de resultados. Es rápida y **no abre la
  ficha del cliente**.
- **Pasada 2 — motivo.** Solo para las líneas que no quedaron activas (o todas
  con `-HistorialSiempre`): abre la ficha, pulsa `Ctrl+Shift+H` y lee el
  HISTORIAL.

### Solo lectura

**El script nunca guarda nada en AC.**

AC pide diligenciar y guardar un formulario **"Solicitud Tickler"** para poder
cerrar una ficha de cliente, y no ofrece forma de cancelarlo: si se cierra el
Tickler sin guardar, la ficha simplemente no se cierra. Guardar ese tickler
dejaría un registro en producción por cada número consultado.

Por eso, cuando la automatización necesita cerrar una ficha, **reinicia AC**
en vez de guardar el tickler. Ese es el motivo de que la pasada 2 tarde
~45 segundos por línea, y de que la pasada 1 evite abrir fichas.

### Velocidad y paralelismo

Tiempos medidos sobre este equipo (i5-4590T, 8 GB):

| Instancias | Por número | Proyección 3000 registros |
|---|---|---|
| 1 | 5,4 s | ~4,5 h |
| 2 | 3,5 s | ~2,9 h |

El paralelismo **no escala de forma lineal** (2 instancias dan 1,54×, no 2×).
El motivo está abajo: los clics tienen que turnarse el primer plano.

Antes de subir de 4 instancias conviene consultarlo con quien administra AC:
cada instancia es una sesión completa contra el Oracle de producción.

### Por qué ocupa el equipo

La automatización no mueve el ratón: envía los clics como mensajes de Windows
directamente a cada control. Pero hay un detalle medido que no se puede
esquivar: **AC solo atiende esos clics si esa instancia es la ventana en primer
plano en ese instante**, y además el primer clic sobre el panel actúa de
"cebador" (deja el foco dentro) — el que cuenta es el siguiente.

Por eso las ventanas de AC saltan al frente mientras corre y **no conviene usar
el equipo al mismo tiempo**. Lo que sí funciona con la ventana de fondo es la
espera de Oracle y la lectura de las grillas, que es la mayor parte del tiempo:
de ahí que varias instancias puedan solaparse aunque se turnen el primer plano.

### Detalles técnicos

- AC es una aplicación VB6 (`ThunderRT6`). Sus grillas
  (`ListView20WndClass`) **no exponen contenido por UI Automation**: se leen
  con mensajes `LVM_*` y memoria reservada dentro del proceso de AC, así que
  los datos salen como texto exacto y **no por OCR**.
- El sondeo usa `LVM_GETITEMCOUNT` (un solo mensaje). Leer la grilla completa
  en cada ciclo satura el bucle de mensajes de AC hasta el punto de que deja
  de procesar los clics que se le envían.
- El número se escribe con `WM_SETTEXT`, no con pulsaciones: al hacer clic en
  el panel de criterios AC desvía el foco del teclado a un PictureBox.
- `Ctrl+Shift+H` se envía con `keybd_event`; `SendKeys` no funciona con estos
  formularios.
- El botón **Consultar** de la ventana de resultados es el único control que
  no responde a mensajes y exige el ratón real. Por eso la pasada del
  historial usa una sola instancia y no se puede paralelizar.
- Poner una ventana en primer plano requiere recuperar el permiso que Windows
  concede solo a quien generó la última entrada del usuario. Se hace
  enganchándose a la cola de entrada del hilo con el foco; el truco de pulsar
  ALT queda como último recurso porque puede activarle la barra de menú a otra
  instancia de AC y dejarla colgada.
- Los handles de ventana se descubren en cada ejecución por clase y título,
  nunca están fijos en el código.
- La contraseña se verifica **dentro del formulario antes de pulsar Aceptar**.
  El dominio bloquea la cuenta tras varios intentos fallidos, así que el script
  no gasta intentos: si el campo no quedó bien escrito, aborta sin enviar.

---

## Si algo falla

| Síntoma | Qué hacer |
|---|---|
| `ERROR. La clave de RED no es correcta.` | La clave cambió. Corre `.\Automatizar-Ventas.ps1 -GuardarClave` de nuevo. Ojo: el dominio bloquea la cuenta temporalmente tras varios intentos fallidos. |
| Muchos `NO ENCONTRADO` seguidos | Revisa que la columna del Excel sea la correcta (`-Columna`) y que AC responda a mano. |
| `ACTIVA = REVISAR` | Falló la consulta de ese número. Queda una captura en `salida\capturas_<fecha>\`. Se puede reprocesar solo ese con `-Numeros`. |
| El script se queda pegado | Alguien movió el mouse o tecleó durante la corrida. Ciérralo con `Ctrl+C`, cierra AC y vuelve a lanzarlo. |

---

## Archivos

```
Automatizar-Ventas.ps1   Script principal
Ejecutar.bat             Lanzador de doble clic
lib\Win32.ps1            Interoperabilidad con la API de Windows
lib\Remote.ps1           Lectura de las grillas de AC (otro proceso)
lib\AC.ps1               Flujo de AC: login, búsqueda, ficha, historial
lib\ACParalelo.ps1       Motor de consulta con varias instancias a la vez
lib\Excel.ps1            Lectura del .xlsx sin Excel instalado
entrada\                 Aquí va el Excel de ventas
salida\                  Reportes CSV y capturas de evidencia
config\clave.dat         Clave de red cifrada (se crea con -GuardarClave)
```

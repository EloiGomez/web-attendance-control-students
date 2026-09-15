[English](README.md) | [Català](README.ca.md) | **Español**

# Control de Asistencia

Un sistema ligero de control de asistencia escolar hecho con **Google Apps Script**, creado para sustituir un Excel lento e incómodo con decenas de pestañas que usaba todo un colegio de primaria para anotar las ausencias diarias de los alumnos.

> Construido como proyecto de programación en pareja con IA junto a [Claude](https://claude.com/claude-code): yo dirigí los requisitos, las decisiones de diseño y las pruebas (incluyendo detectar un fallo de caché, un fallo de error silencioso y un problema de rendimiento con mucha carga de datos), mientras que Claude se encargó de la implementación.

## El problema

El flujo de trabajo original era un único archivo Excel con una rejilla enorme por cada clase (una fila por alumno, una columna por cada día lectivo del año — unas 240 filas × 390 columnas por clase), duplicada en decenas de pestañas para todo el colegio. Los profesores tenían que escribir códigos de una letra a mano en cada celda, el archivo tardaba mucho en abrirse y editarse, y no había ninguna forma automática de ver el porcentaje de ausencias de un alumno durante el curso.

## La solución

Una pequeña aplicación web servida directamente desde Google Apps Script, con una Google Sheet detrás que solo hace de base de datos (los profesores nunca la abren directamente):

- **Pasar lista por clase y día**: eliges una clase y una fecha, ves a todos los alumnos en una rejilla de casillas (mañana / tarde / justificada / retraso), y guardas todo el día de una vez.
- **Resumen de ausencias en tiempo real**: porcentaje de asistencia automático por alumno (ponderado — faltar una mañana no cuenta igual que faltar una tarde, siguiendo el horario real del colegio), desglosado entre falta justificada/no justificada, con un nivel de alerta progresivo por colores, un desglose por día de la semana, un buscador y columnas ordenables.
- **Control de acceso**: restringido a una lista explícita de emails del profesorado autorizado, comprobada en el servidor en cada petición — no solo un enlace difícil de adivinar.
- **Herramientas de administración**: una herramienta de un solo clic para "promocionar todas las clases" a final de curso (con una hoja de mapeo de clases editable y una casilla de repetidor por alumno), además de generadores de datos de prueba para hacer pruebas de carga con cientos de alumnos.
- **Rendimiento**: el backend agrupa todas las lecturas de la Sheet, reutiliza un único objeto `Spreadsheet` por ejecución en vez de volver a pedirlo repetidamente, y guarda en caché el resumen calculado (comprimido con gzip para que quepa dentro del límite de tamaño del caché de Apps Script), de forma que volver a consultarlo es casi instantáneo hasta que los datos cambian de verdad.
- Modo oscuro automático (`prefers-color-scheme`), un diseño adaptado a móvil con cabeceras/columnas fijas al hacer scroll, y ninguna dependencia externa — HTML/CSS/JS puro servido a través de `HtmlService`.

## Tecnología

- **Backend**: Google Apps Script (JavaScript, motor V8) — `SpreadsheetApp`, `HtmlService`, `CacheService`, `PropertiesService`.
- **Frontend**: HTML/CSS/JS puro, sin frameworks ni paso de compilación — se comunica con el backend vía `google.script.run`.
- **Almacenamiento de datos**: una Google Sheet, usada como una base de datos estructurada sencilla, no como interfaz.

## Estado

Prototipo funcional, probado con datos sintéticos (cientos de alumnos, miles de registros de asistencia). No contiene ningún dato real de alumnos — la hoja `Students` trae solo nombres de ejemplo.

## Instalación y despliegue

No hace falta ningún entorno local ni paso de compilación — todo funciona dentro de la infraestructura de Google.

1. Crea una nueva Google Sheet en blanco en [sheets.google.com](https://sheets.google.com).
2. Abre **Extensiones → Apps Script**. Esto crea un proyecto de Apps Script vinculado a esa Sheet.
3. En el archivo `Code.gs` por defecto, borra el contenido de ejemplo y pega el [`Code.gs`](Code.gs) de este repositorio.
4. Añade un archivo nuevo de tipo **HTML** (el `+` junto a "Archivos"), llámalo exactamente `Index` (sin extensión), y pega el [`Index.html`](Index.html) de este repositorio.
5. Guarda el proyecto.
6. En el desplegable de funciones de arriba, elige **`setup`** y pulsa **Ejecutar**. La primera vez te pedirá autorizar el script — acéptalo. Esto crea las hojas `Config`, `Students`, `Teachers`, `Records` y `Promotion` con datos de ejemplo y valores por defecto razonables.
7. (Opcional) Ejecuta **`generateTestStudents`** y **`generateTestRecords`** para cargar unos cientos de alumnos y registros de asistencia sintéticos, útil para probar el resumen y la ordenación/buscador sin escribir nada a mano.
8. Añade los emails de quien deba tener acceso a la hoja **`Teachers`** — si esa hoja está vacía, la aplicación deja entrar a cualquiera, así que este paso importa antes de compartir el enlace.
9. Ve a **Implementar → Nueva implementación**, elige el tipo **Aplicación web**. Configura:
   - **Ejecutar como**: *Usuario que accede a la aplicación web* — así la identidad de Google de cada visitante es la que usa de verdad el control de acceso (y la columna de auditoría "Updated By").
   - **Quién tiene acceso**: *Cualquier usuario con cuenta de Google* (o *Cualquier usuario dentro de [dominio]*, si se despliega desde una organización de Google Workspace).
10. Pulsa **Implementar**, y abre la URL `.../exec` resultante. La primera vez que cada usuario la abra, Google le pedirá autorizar la aplicación con su propia cuenta — es normal, ya que se ejecuta como cada visitante, no como el desarrollador.

Para publicar un cambio de código más adelante: edita los archivos, guarda, y luego **Implementar → Gestionar implementaciones → ✏️ → Versión: Nueva versión → Implementar** — la misma URL `.../exec` sigue funcionando, solo empieza a servir el código actualizado.

## Archivos

- [`Code.gs`](Code.gs) — lógica del servidor (rutas, control de acceso, acceso a datos, lógica de negocio, herramientas de administración).
- [`Index.html`](Index.html) — el frontend de una sola página.

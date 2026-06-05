# 📄 Generador de Documentos

Aplicación de escritorio para generar contratos, finiquitos y cartas automáticamente a partir de plantillas Word y datos desde Excel.

---

## Funcionalidades

- Lee cualquier Excel automáticamente y detecta sus columnas
- Soporta múltiples plantillas Word (contratos, finiquitos, cartas, etc.)
- Buscador de personas por nombre o RUT
- Detecta campos del Word que no están en el Excel y los pide manualmente
- Genera un archivo Word nuevo sin modificar la plantilla original
- Se actualiza automáticamente cuando hay una nueva versión

---

## Instalación para usuarios finales

Descarga el archivo `Generador de Documentos.exe` desde la sección [Releases](../../releases/latest) y ejecútalo con doble clic. No requiere instalar nada más.

La aplicación se actualiza sola cuando hay una nueva versión disponible.

---

## Instalación para desarrollo

**Requisitos:** Python 3.10 o superior

```bash
# 1. Clonar el repositorio
git clone https://github.com/mvillarios/docsGen.git
cd docsGen

# 2. Crear entorno virtual
python -m venv venv

# 3. Activar el entorno virtual
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate

# 4. Instalar dependencias
pip install customtkinter openpyxl python-docx pillow pyinstaller

# 5. Ejecutar la aplicación
python app.py
```

---

## Compilar el ejecutable manualmente

```bash
pyinstaller --onefile --windowed \
  --add-data "version.txt;." \
  --name "Generador de Documentos" \
  app.py
```

El `.exe` quedará en la carpeta `dist/`.

---

## Flujo de releases (CI/CD)

Este repositorio usa GitHub Actions para compilar y publicar automáticamente.

```
Push a dev
    ↓
GitHub Actions compila el .exe en Windows
    ↓
Merge automático dev → main
    ↓
Se crea un Release con el .exe adjunto
    ↓
La app detecta la nueva versión y se actualiza sola
```

**Para publicar una nueva versión:**

1. Edita `version.txt` y aumenta el número (ej: `1.0.0` → `1.0.1`)
2. Haz tus cambios en el código
3. Sube a `dev`:

```bash
git add .
git commit -m "descripción del cambio"
git push origin dev
```

El resto es automático.

---

## Cómo preparar las plantillas Word

En los documentos Word, reemplaza los valores variables con marcadores entre llaves dobles. El nombre del marcador debe coincidir exactamente con el nombre de la columna en el Excel.

| En vez de escribir  | Escribe              |
|---------------------|----------------------|
| Juan Pérez          | `{{nombre}}`         |
| 12.345.678-9        | `{{rut}}`            |
| Calle Los Pinos 123 | `{{direccion}}`      |
| Operario            | `{{cargo}}`          |
| Centro Norte        | `{{centro_de_costo}}`|
| 01/01/2025          | `{{fecha}}`          |

**Ejemplo de párrafo en el Word:**

> Por medio de la presente, se informa que el trabajador `{{nombre}}`, RUT `{{rut}}`, con domicilio en `{{direccion}}`, ha sido contratado como `{{cargo}}` adscrito al centro de costo `{{centro_de_costo}}`.

Si un campo del Word no existe en el Excel (como `{{fecha}}`), la aplicación lo detecta automáticamente y lo pide manualmente antes de generar.

---

## Cómo preparar el Excel

- La **primera fila** debe contener los nombres de las columnas
- Cada fila siguiente corresponde a una persona
- Los nombres de columna deben coincidir con los marcadores usados en los Word

**Ejemplo:**

| nombre           | rut           | direccion              | cargo    | centro_de_costo | sueldo  |
|------------------|---------------|------------------------|----------|-----------------|---------|
| María González   | 15.234.567-8  | Av. Providencia 1234   | Analista | Finanzas        | 850000  |
| Carlos Rodríguez | 12.876.543-2  | Los Pinos 456          | Técnico  | Operaciones     | 720000  |

---

## Estructura del proyecto

```
├── .github/
│   └── workflows/
│       └── release.yml      # Pipeline de CI/CD
├── app.py                   # Aplicación principal
├── updater.py               # Módulo de auto-actualización (independiente)
├── version.txt              # Versión actual (ej: 1.0.0)
├── .gitignore
└── README.md
```

---

## Preguntas frecuentes

**¿El Word original se modifica al generar un documento?**
No. Siempre se crea un archivo nuevo; la plantilla original queda intacta.

**¿Puedo tener varias plantillas Word?**
Sí. En cada uso se selecciona la plantilla que se necesite.

**¿Qué pasa si no hay conexión a internet al abrir la app?**
La verificación de actualizaciones falla silenciosamente y la app funciona con normalidad.

**¿El repositorio puede ser privado?**
Sí, pero requiere agregar un token de GitHub en `updater.py` para que la app pueda consultar los releases.

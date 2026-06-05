# 📄 Generador de Documentos
## Cómo usar la aplicación

---

## Instalación (una sola vez)

1. Instala Python desde https://www.python.org/downloads/
   - Marca la casilla **"Add Python to PATH"** durante la instalación

2. Abre la carpeta de la aplicación, haz clic derecho en un espacio vacío
   y selecciona **"Abrir en Terminal"** (o "PowerShell")

3. Copia y pega este comando, luego presiona Enter:
   ```
   pip install customtkinter openpyxl python-docx pillow
   ```

4. Cuando termine, escribe:
   ```
   python app.py
   ```

---

## Cómo preparar las plantillas Word

En tus documentos Word, reemplaza los datos variables con marcadores así:

| En vez de escribir | Escribe |
|--------------------|---------|
| Juan Pérez         | {{nombre}} |
| 12.345.678-9       | {{rut}} |
| Calle Los Pinos 123| {{direccion}} |
| Operario           | {{cargo}} |
| Centro Norte       | {{centro_de_costo}} |
| 01/01/2025         | {{fecha}} |

**El nombre del marcador debe coincidir con el nombre de la columna en tu Excel.**

Ejemplo de párrafo en el Word:
> "Por medio de la presente, se informa que el trabajador **{{nombre}}**, RUT **{{rut}}**,
> con domicilio en **{{direccion}}**, ha sido contratado como **{{cargo}}**..."

---

## Cómo preparar el Excel

- La **primera fila** debe tener los nombres de las columnas (ej: nombre, rut, cargo)
- Cada fila siguiente = una persona
- No mezclar datos en la primera fila

---

## Pasos en la aplicación

1. **Cargar archivos**: selecciona tu Excel y el Word base
2. **Seleccionar persona**: busca por nombre o RUT
3. **Revisar datos**: corrige o agrega campos extras (como la fecha)
4. **Generar**: elige dónde guardar el documento nuevo

---

## Preguntas frecuentes

**¿El Word original se modifica?**
No. Siempre se crea un archivo nuevo; el original queda intacto.

**¿Qué pasa si un campo del Word no está en el Excel?**
La app lo detecta y te pide que lo ingreses manualmente en el paso 3.

**¿Puedo tener varios templates Word?**
Sí. Simplemente selecciona el que necesites en cada uso.

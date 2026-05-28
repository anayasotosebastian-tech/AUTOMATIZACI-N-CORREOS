# Portal de Planes de Lealtad – Grupo Farmatodo

Aplicación web para generar convenios de lealtad, exportarlos a PDF y preparar correos.

## Instalación

```bash
cd app
pip install -r requirements.txt
```

## Uso

```bash
cd app
python3 app.py
```

Abre `http://localhost:5000` en tu navegador.

## Vista previa PDF

La vista previa usa **LibreOffice** (modo headless) para convertir el Word a PDF.  
Si no está instalado, descarga el `.docx` para revisarlo.

### Instalar LibreOffice (Ubuntu/Debian)

```bash
sudo apt install libreoffice -y
```

### Instalar LibreOffice (macOS)

```bash
brew install --cask libreoffice
```

## Archivos requeridos en `app/`

| Archivo | Descripción |
|---|---|
| `plantilla.docx` | Convenio base con campos resaltados en amarillo |
| `materiales.xlsx` | Catálogo de SKUs y laboratorios (columnas: SAP, SKU, DESCRIPCION, ID LAB, LAB) |

## Flujo

1. Llena los datos del formulario
2. Pega los SKUs y haz clic en **Buscar productos** (SAP y nombre se completan automáticamente)
3. Haz clic en **Vista previa** — puedes modificar cualquier celda antes de continuar
4. Haz clic en **Aprobar y preparar correo**
5. Revisa/edita el correo y haz clic en **Abrir en Outlook/correo**

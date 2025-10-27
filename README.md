# 📋 Sitio para el Alistamiento de Dispositivos

Aplicación web desarrollada en Flask como principal checklists de alistamiento de equipos Proquinal y Comercializadora Calypso.

## 🚀 Características

- ✅ 5 tipos diferentes de checklists
- 📱 Interfaz responsiva y moderna con Bootstrap 5
- 📊 Generación automática de Excel basado en plantillas
- 🎨 Diseño intuitivo con íconos y emojis
- 📝 Validación de formularios en tiempo real
- 📈 Barra de progreso para seguimiento

## 💄 Apariencia de la herramienta

![Captura de pantalla: home ](/static/preview_1.png)

![Captura de pantalla: datos ](/static/preview_2.png)

![Captura de pantalla: formulario ](/static/preview_3.png)

## 📦 Instalación

### 1. Clonar o descargar el proyecto

```bash
git clone https://github.com/josuerom/app-checklist-proquinal.git
cd app-checklist-proquinal
```

### 2. Crear entorno virtual (recomendado)

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

### 3. Instalar dependencias del proyecto

```bash
pip install -r requirements.txt
```

**Nombres de archivos requeridos:**
- `plantilla_pc.xlsx`
- `plantilla_terminales.xlsx`
- `plantilla_macos.xlsx`
- `plantilla_tablets.xlsx`
- `plantilla_calypso.xlsx`

**Formato de las plantillas:**
- Columna A: ID de pregunta (números 1, 2, 3, etc.)
- Columna B: Texto de la pregunta (opcional)
- Columna C: Respuesta (aquí se escribirán los resultados)

## ▶️ Ejecución

```bash
python app.py
```

La aplicación estará disponible en: `http://localhost:9015`

## 📁 Estructura del Proyecto

```
app-checklist_pqn/
├── app.py                      # Aplicación Flask principal
├── requirements.txt            # Dependencias
├── README.md                   # Documentación del proyecto
├── static/
│   └── img/
│       ├── stefanini.png      # Logo Stefanini
│       └── proquinal.png       # Logo Proquinal
├── templates/
│   ├── index.html             # Página principal
│   ├── checklist_form.html    # Formulario de datos
│   └── checklist.html         # Checklist interactivo
└── templates_excel/           # Archivos plantillas
    ├── plantilla_pc.xlsx
    ├── plantilla_terminales.xlsx
    ├── plantilla_macos.xlsx
    ├── plantilla_tablets.xlsx
    └── plantilla_calypso.xlsx
```

## 🎯 Flujo de Uso

1. **Seleccionar tipo de checklist** en la página principal
2. **Ingresar datos iniciales:**
   - Número de activo fijo
   - Propietario del dispositivo
   - Cargo del propietario
3. **Completar el checklist** marcando OK o N/A para cada pregunta
4. **Generar Excel** al finalizar todas las preguntas
5. **Descargar** el archivo generado automáticamente

## 📊 Tipos de Checklist Disponibles

| Tipo | Empresa | Dispositivo | Preguntas |
|------|---------|-------------|-----------|
| PC | Proquinal | Equipos de escritorio y portátiles | 10 |
| Terminales | Proquinal | Terminales punto de venta | 8 |
| MacOS | Proquinal | MacBook e iMac | 10 |
| Tablets | Proquinal | Tablets corporativas | 9 |
| Calypso | Calypso | Sistemas CCS CBQ | 10 |

## 🔧 Personalización

### Agregar un nuevo tipo de checklist

Edita el diccionario `CHECKLISTS` en `app.py`:

```python
'nuevo_tipo': {
    'titulo': 'Checklist Nuevo Tipo 2025',
    'empresa': 'Proquinal',
    'tipo': 'TipoDispositivo',
    'plantilla': 'plantilla_<impresoras>.xlsx',
    'preguntas': [
        'Pregunta 1',
        'Pregunta 2',
        # ... más preguntas
    ]
}
```

### Modificar preguntas existentes

Edita la lista `preguntas` en el checklist correspondiente dentro del diccionario `CHECKLISTS`.

## 📝 Formato del Archivo Excel Generado

Nombre del archivo de salida:
```
Activo <activo_fijo> Checklist <empresa> <tipo> <propietario> <cargo>.xlsx
```

Ejemplo:
```
Activo 35094 Checklist Proquinal PC Pepa_la_cerdita Analista_Sistemas.xlsx
```

## 🔒 Seguridad

- Las sesiones están protegidas con secret_key
- Los datos solo se almacenan durante la sesión actual
- No se guarda información personal en la base de datos
- Los archivos generados se almacenan temporalmente en `/output/`

## 🚢 Despliegue en Producción

### Con Gunicorn

```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:9015 app:app
```

### Con Docker (opcional)

```dockerfile
FROM python:3
WORKDIR /app
COPY . /app
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 9015
RUN python app.py
```

## 🐛 Solución de Problemas

### Error: "No module named 'openpyxl'"
```bash
pip install openpyxl
```

### Error: "Template not found"
Verifica que la carpeta `templates/` esté en el directorio raíz del proyecto.

### Los logos no aparecen
Asegúrate de que los archivos estén en `static/img/` y tengan los nombres correctos.

### El Excel no se genera
- Verifica que la plantilla exista en `templates_excel/`
- Asegúrate de que la columna A contenga los números de pregunta
- Revisa que el archivo tenga permisos de escritura

## 📧 Soporte

Para preguntas o problemas:
- 📧 Email: josue.romero@spradling.group
- 📞 Teléfono: +57 (310) 864 3149

## 📄 Licencia

©2025 JRJ Bogotá D.C - Stefanini Group CO. Todos los derechos reservados.

---

**Desarrollado por el loco empedernido Josué**

# CIARP - Sistema de Notificaciones por Correo

Sistema automatizado para notificar a docentes de la **Universidad del Quindío** sobre los puntos salariales aprobados en el **Comité Interno de Asignación y Reconocimiento de Puntaje (CIARP)**.

## ¿Qué hace?

Lee el archivo Excel del CIARP (con hojas de Títulos, Artículos, Libros de Ensayo, Libros de Texto y Premios), filtra los docentes con puntaje mayor a 0, y envía un correo HTML personalizado a cada uno via **Gmail API (OAuth 2.0)**.

## Archivos principales

| Archivo | Descripción |
|---|---|
| `send_gmail.ps1` | Script principal de envío de notificaciones |
| `generar_reporte_ciarp.ps1` | Genera reporte completo en Excel y Word |
| `extract_data.ps1` | Extrae y parsea datos del Excel CIARP |

## Modos de ejecución

```powershell
# Prueba: envía 1 correo de ejemplo a jhvargas
powershell -File send_gmail.ps1 -Test

# Prueba total: envía TODOS los correos a jhvargas para revisión
powershell -File send_gmail.ps1 -TestAll

# Real: envía correos a cada docente
powershell -File send_gmail.ps1 -Real
```

## Configuración previa

1. Crear proyecto en [Google Cloud Console](https://console.cloud.google.com/)
2. Habilitar **Gmail API**
3. Crear credenciales OAuth 2.0 → descargar como `credentials.json`
4. La primera ejecución abrirá el navegador para autenticarse

## Archivos necesarios (NO incluidos en el repo)

- `credentials.json` — Credenciales OAuth de Google Cloud
- `tokens.json` — Token de acceso (se genera automáticamente)
- `ciarp 2.xlsx` — Archivo Excel del CIARP con las hojas de productos
- `correos.xlsx` — Archivo Excel con correos de docentes de planta

## Características

- ✅ Manejo de **celdas combinadas** en Excel
- ✅ Agrupa múltiples productos por docente en un solo correo
- ✅ Excluye automáticamente productos con puntaje 0 o vacío
- ✅ Nota especial para docentes en **comisión académica-administrativa**
- ✅ Genera reporte Excel con hoja de Notificados y No Notificados
- ✅ Genera informe Word formal como acta
- ✅ Correo HTML con diseño institucional

## Estructura del correo

El correo incluye:
- Nombre y programa del docente
- Tabla con cada producto, concepto y puntaje
- Puntaje total aprobado
- (Cuando aplica) Nota sobre comisión académica-administrativa
- Firma de la Jefe de la Oficina de Asuntos Profesorales

---
**Universidad del Quindío — Oficina de Asuntos Profesorales**

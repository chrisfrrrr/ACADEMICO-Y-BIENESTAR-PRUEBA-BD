# AVE Seguimiento PRO v12

Sistema de seguimiento y derivación académica para asesores AVE.

## Novedades v12

- Mantiene el flujo de Excel manual/local o sincronizado.
- Agrega modo **Supabase en línea**, recomendado para Streamlit Cloud.
- Permite leer y guardar la base institucional sin rutas locales, sin Azure y sin Microsoft Graph.
- Conserva conexión con Canvas, selector de cursos, clasificación de riesgo, mensajes y derivaciones por asesor de bienestar.
- Permite exportar la base a Excel como respaldo institucional.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Configuración Supabase

1. Crear proyecto en Supabase.
2. Ejecutar `supabase_schema.sql` en SQL Editor.
3. Guardar credenciales en Streamlit Cloud > Settings > Secrets:

```toml
[supabase]
url = "https://TU-PROYECTO.supabase.co"
service_role_key = "TU_SERVICE_ROLE_KEY"
```

4. En la app, ir a **Base de datos > Usar Supabase en línea**.

## Estructura de datos

La base se guarda en una tabla llamada `ave_base_datos`, donde cada fila representa una hoja lógica:

- Estudiantes
- Asesores_Bienestar
- Historial_Estudiantes
- Derivaciones
- Mensajes_Enviados
- Consultas_Canvas
- Configuracion

Este diseño evita depender de un Excel físico en Streamlit Cloud y permite exportar todo nuevamente a Excel cuando se necesite.

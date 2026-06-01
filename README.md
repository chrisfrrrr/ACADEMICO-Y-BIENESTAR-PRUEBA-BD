# AVE Seguimiento PRO v13 - Multi-curso y multi-asesor

Esta versión permite trabajar con una base general institucional, separando cada análisis por periodo/cohorte, curso, ID de curso Canvas, sección, ID de sección Canvas, semana y asesor académico.

## Novedades v13

- Base general para varios cursos, secciones y asesores académicos.
- No mezcla el historial de un mismo estudiante entre cursos diferentes.
- Dashboard con filtros por periodo, curso, sección, asesor académico, asesor de bienestar y riesgo.
- Nueva pestaña **Institucional** para ver resumen por curso/sección/asesor.
- Derivaciones contextualizadas con periodo, curso y sección.
- Lectura más robusta de CSV institucionales con separador coma o punto y coma.
- Compatible con Supabase como base en línea para Streamlit Cloud.

## Uso recomendado

1. Crear o leer base desde Supabase.
2. Completar en la barra lateral: periodo/cohorte, semana, curso manual y sección si aplica.
3. Validar token de Canvas y cargar cursos.
4. Seleccionar curso y ejecutar análisis.
5. Revisar Dashboard e Institucional.
6. Generar mensajes y derivaciones.
7. Guardar cambios en Supabase o exportar Excel.

## Supabase

Ejecutar `supabase_schema.sql` en el SQL Editor del proyecto. Luego configurar en Streamlit Cloud > Settings > Secrets:

```toml
[supabase]
url = "https://TU-PROYECTO.supabase.co"
service_role_key = "TU_SERVICE_ROLE_KEY"
```

La tabla compacta utilizada es `ave_base_datos`, donde cada hoja se guarda como JSONB.


## Mejora integrada: derivaciones contextualizadas

La pestaña **Derivaciones** ahora genera paquetes organizados por asesor de bienestar, curso/sección, asesor académico remitente y nivel de riesgo. Cada documento de derivación incluye el contexto académico completo del caso: periodo, curso, ID Canvas del curso, sección, ID Canvas de sección, semana de análisis, asesor académico remitente, asesor de bienestar asignado y nivel de prioridad.

El ZIP generado incluye un listado general de todas las derivaciones y un listado específico dentro de la carpeta de cada asesor de bienestar.


## Versión 14 - Supabase multi-tabla

Esta versión guarda la información en tablas separadas de Supabase: estudiantes, asesores_bienestar, historial_estudiantes, derivaciones, mensajes_enviados, consultas_canvas y configuracion. La tabla ave_base_datos queda como respaldo. Para configurar la conexión, revisar INSTRUCCIONES_V14_SUPABASE_MULTITABLA.txt.

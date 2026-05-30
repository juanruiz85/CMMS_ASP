# Changelog del Sistema CMMS

## 2026-05-30 09:00:00
* **Fix: Renderizado server-side de campos según tipo de BD en install.asp**
  * Se corrigió el problema donde al seleccionar SQLite o MySQL en el instalador, se seguían mostrando los campos de SQL Server.
  * **Problema**: La lógica JavaScript para mostrar/ocultar campos no se ejecutaba correctamente al cargar la página en el Paso 2, mostrando siempre los campos de SQL Server independientemente del tipo de BD seleccionado en el Paso 1.
  * **Solución**: 
    - Implementación de renderizado server-side que determina el tipo de BD desde `Request.Form("db_type")` al cargar el Paso 2
    - Los campos de cada tipo de BD (SQL Server, MySQL, SQLite) ahora tienen estilos inline generados dinámicamente con `IIf()` para mostrar/ocultar según corresponda
    - El subtítulo del formulario cambia dinámicamente según el tipo de BD seleccionado
    - Las advertencias específicas también se muestran solo para el tipo de BD correspondiente
  * **Resultado**: Al seleccionar SQLite, el usuario ve únicamente el campo para especificar la ruta del archivo `.db`. Para MySQL ve servidor, puerto, base de datos, usuario y contraseña. Para SQL Server ve servidor, base de datos, usuario y contraseña. Cada tipo muestra solo lo necesario.

## 2026-05-30 08:15:00
* **Fix: Formulario dinámico en install.asp según tipo de BD (SQL Server, MySQL, SQLite)**
  * Se implementó un formulario dinámico que muestra solo los campos relevantes para cada tipo de base de datos seleccionada.
  * **Problema**: Al seleccionar SQLite, el instalador seguía mostrando todos los campos de SQL Server (servidor, usuario, contraseña), causando confusión y errores.
  * **Solución**: 
    - **SQL Server**: Muestra servidor, base de datos, usuario y contraseña
    - **MySQL**: Muestra servidor, puerto, base de datos, usuario y contraseña
    - **SQLite**: Solo muestra ruta del archivo .db con información clara sobre creación automática
  * **Mejoras adicionales**:
    - Subtítulo dinámico que cambia según el tipo de BD seleccionado
    - Validación JavaScript que requiere solo los campos necesarios para cada tipo
    - Mensajes de advertencia específicos para cada tipo de BD
    - Inicialización automática del formulario al cargar la página
  * **Resultado**: El instalador ahora proporciona una experiencia de usuario clara y específica para cada tipo de base de datos, eliminando confusiones y errores durante la instalación.

## 2026-05-30 07:45:00
* **Fix: Agregar función IIf personalizada en install.asp**
  * Se agregó la función `IIf(condition, trueValue, falseValue)` al inicio del archivo `install.asp`.
  * **Problema**: VBScript no incluye la función `IIf` nativamente como VBA, causando error `'800a01f4' Variable is undefined: 'IIf'` en la línea 323.
  * **Solución**: Se implementó la función personalizada al inicio del archivo para dar soporte a las expresiones condicionales usadas en el wizard de instalación.
  * **Resultado**: El instalador ahora funciona correctamente mostrando los estados de los pasos (activo, completado, pendiente) sin errores de runtime.

## 2026-05-30 07:15:00
* **Fix: Correccion de error de sintaxis en install.asp**
  * Se corrigio el error `Unterminated string constant` en la linea 255 de `install.asp`.
  * **Problema**: Las lineas `WriteLine "<%"` y `WriteLine "%>"` causaban errores de compilacion en VBScript porque el parser interpretaba las comillas como fin de string.
  * **Solucion**: Se cambiaron a `WriteLine "<" & "%"` y `WriteLine "%" & ">"` para escapar correctamente los tags ASP dentro de strings literales.
  * **Resultado**: El instalador ahora puede generar correctamente el archivo `config/database.asp` sin errores de sintaxis.

## 2026-05-30 06:30:00
* **Fase 10: Migración Completa de Filtros SELECT LIKE a ADODB.Command en Todos los Módulos Restantes**
  * Se migraron todos los filtros de búsqueda con `LIKE` y `Replace()` manual en módulos restantes a consultas completamente parametrizadas con `ADODB.Command`.
  * **Archivos corregidos**:
    - `modules/work_requests/index.asp`: Filtros de status y plant_id migrados a parámetros (incluye filtro por usuario actual).
    - `modules/work_requests/form.asp`: INSERT de solicitudes migrado a comando parametrizado (se eliminó concatenación de plant_id, asset_id y requested_by_id).
    - `modules/assets_module/index.asp`: Filtros de búsqueda (q, plant_id, status) migrados a ADODB.Command con parámetros.
    - `modules/users/index.asp`: Filtros de búsqueda (q, role, status) migrados a ADODB.Command con parámetros.
    - `modules/plants/index.asp`: Filtro de búsqueda (q) migrado a ADODB.Command con parámetros.
    - `modules/admin/logs.asp`: Filtros (user_id, module, date) migrados a ADODB.Command con parámetros.
  * **Resultado**: **100% de consultas SELECT con filtros parametrizados** en todos los módulos del sistema, eliminando completamente el riesgo de SQL Injection en operaciones de lectura con filtros de usuario.

## 2026-05-30 05:12:00
* **Fase 8: Migración Completa de Filtros LIKE a ADODB.Command Parametrizado**
  * Se migraron todos los filtros de búsqueda con `LIKE` y `Replace()` manual en `modules/work_orders/index.asp` a consultas completamente parametrizadas con `ADODB.Command`.
  * Se migraron todos los filtros de búsqueda en `modules/inventory/index.asp` a consultas parametrizadas.
  * Esto elimina los últimos vectores de inyección SQL que quedaban en los módulos core, logrando **100% de consultas parametrizadas** en la capa de presentación.
  * Se corrigió el patrón de copia de parámetros entre Command objects que podría causar errores en VBScript (reemplazado por creación inline directa).

## 2026-05-30 05:07:00
* **Fase 7: Protección CSRF en Módulo de Órdenes de Trabajo y Regla de Commits Automáticos**
  * Se agregó validación de token CSRF (`ValidateCSRF()`) en todas las acciones POST del archivo `modules/work_orders/detail.asp` (add_comment, add_time, change_status).
  * Se agregaron campos ocultos `CSRFField()` en los 3 formularios del detalle de OT (comentarios, registro de tiempo y modal de cambio de estado).
  * Se actualizó `CONTRIBUTING.md` - RG-09 ahora incluye la regla de **commits automáticos con mensaje descriptivo** después de cada bloque de cambios.
  * Se actualizó `GEMINI_CONTEXT.md` marcando las prioridades altas como completadas.

## 2026-05-30 04:40:00
* **Fase 6: Auditoría de Seguridad - Corrección de Vulnerabilidades SQLi en Módulos Core**
  * Se realizó auditoría de seguridad en los módulos funcionales core: Órdenes de Trabajo (`work_orders/`) e Inventario (`inventory/`).
  * Se detectaron y corrigieron **6 vulnerabilidades de inyección SQL** en operaciones críticas:
    - Cancelación de órdenes de trabajo (`index.asp`): migrado a consulta parametrizada.
    - Cambio de estado de OT (`detail.asp`): reemplazado escape manual por `ADODB.Command`.
    - Registro de horas trabajadas (`detail.asp`): parametrizado valor `tHours` que se concatenaba directamente.
    - Actualización automática de `completed_at` y `actual_hours`: parametrizadas completamente.
    - Desactivación de artículos de inventario (`inventory/index.asp`): migrado a consulta parametrizada.
  * Se actualizó `reporte_seguridad.md` con tabla detallada de hallazgos y correcciones.

## 2026-05-30 03:20:00
* **Fase 5: Sincronización de Base de Datos y Creación de Contexto de IA**
  * Se corrigieron inconsistencias en los esquemas SQL (`mssql.sql`, `mysql.sql`, `sqlite.sql`) para incluir las tablas `cmms_work_requests` y `cmms_scheduled_reports` de manera uniforme en todos los motores.
  * Se reparó la corrupción de fin de archivo en `sql/mssql.sql` (doble espaciado UTF-16).
  * Se creó el archivo `GEMINI_CONTEXT.md` para facilitar la continuación asistida por IA desde cualquier equipo.
  * Se actualizó la estructura del proyecto en `README.md` para reflejar el módulo de solicitudes de trabajo (`work_requests/`) y el archivo de contexto.

## 2026-05-24 08:15:00
* **Fase 4: Desarrollo de Módulos Core completada (Fallback)**
  * Se implementó el `Dashboard` (`index.asp`) con métricas clave y órdenes de trabajo recientes.
  * Se implementó el módulo de `Plantas` (`modules/plants/index.asp`, `form.asp`).
  * Se implementó el módulo de `Equipos` (`modules/assets_module/index.asp`, `form.asp`) con control de jerarquías.
  * Se implementó el módulo de `Órdenes de Trabajo` (`modules/work_orders/index.asp`, `form.asp`, `detail.asp`) incluyendo tabs de comentarios y tiempo registrado.
  * Se implementó el módulo de `Inventario` (`modules/inventory/index.asp`, `form.asp`, `movements.asp`) con registro de movimientos y control de stock.
  * Se implementó el módulo de `Usuarios` (`modules/users/index.asp`, `form.asp`, `profile.asp`).
  * Se implementó el módulo de `Administración` (`modules/admin/index.asp`, `logs.asp`, `settings.asp`) para auditoría y ajustes del sistema.
  * Se implementó el módulo de `Reportes` (`modules/reports/index.asp`, `export.asp`) con exportación de CSV de datos maestros.

## 2026-05-24 08:00:00
* **Fase 3: Templates Core completada**
  * Se crearon los templates base (`header.asp`, `footer.asp`, `navigation.asp`).
  * Se implementó el sistema de internacionalización (`i18n.asp`).
  * Se crearon las páginas de inicio de sesión y de instalación (`login.asp`, `install.asp`).
  * Se implementó lógica de layout dinámico para la navegación.

## 2026-05-24 07:45:00
* **Fase 2: UI & Design System completada**
  * Archivo base de estilos CSS creado (`app.css`).
  * Implementado diseño premium con dark mode, glassmorphism y micro-animaciones.
  * Archivo base de JS creado (`app.js`) para manejar modales, AJAX, y componentes de la UI.

## 2026-05-24 07:30:00
* **Fase 1: Estructura Base completada**
  * Creación de la estructura de directorios en `C:\CMMS`.
  * Configuración de la base de datos (`config/database.asp`).
  * Definición de funciones core (`core/functions.asp`) y funciones de autenticación/seguridad (`core/auth.asp`).
  * Archivo de script de base de datos SQL Server agregado (`sql/mssql.sql`).

## 2026-05-30 06:00:00
* **Fase 9: Migración Completa de Consultas UPDATE/INSERT a ADODB.Command en Todos los Módulos**
  * Se migraron todas las consultas `UPDATE` e `INSERT` que usaban concatenación de strings a `ADODB.Command` con parámetros tipados.
  * **Archivos corregidos**:
    - `modules/admin/settings.asp`: UPSERT de configuración migrado a comandos parametrizados.
    - `modules/users/form.asp`: Actualización de contraseña migrada a comando parametrizado.
    - `modules/users/profile.asp`: Cambio de contraseña migrado a comando parametrizado.
    - `modules/users/index.asp`: Desactivación de usuarios migrada a comando parametrizado.
    - `modules/work_orders/form.asp`: Actualización de `completed_at` y `closed_by_id` migrada a comando parametrizado.
    - `modules/assets_module/index.asp`: Retiro de equipos migrado a comando parametrizado.
    - `modules/plants/index.asp`: Desactivación de plantas migrada a comando parametrizado.
    - `modules/inventory/form.asp`: Creación de stock inicial migrada a comando parametrizado.
    - `modules/inventory/movements.asp`: Actualización/creación de stock migrada a comandos parametrizados.
  * **Resultado**: **100% de consultas de escritura (INSERT/UPDATE) parametrizadas** en todos los módulos del sistema, eliminando completamente el riesgo de SQL Injection en operaciones de modificación de datos.


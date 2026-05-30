# Reporte de Auditoría y Mejoras de Seguridad CMMS

Este documento sirve como registro vivo de las correcciones de seguridad aplicadas al sistema.

## 2026-05-30: Auditoría de Seguridad - Corrección de Inyecciones SQL en Módulos Core

### Hallazgos Críticos

Durante la auditoría de seguridad de los módulos funcionales core (Órdenes de Trabajo e Inventario) se identificaron las siguientes vulnerabilidades de **Inyección SQL (SQLi)** que han sido corregidas:

#### 1. Módulo de Órdenes de Trabajo (`modules/work_orders/`)

**a) Acción de cancelación (`index.asp`):**
- **Vulnerabilidad**: El ID de la orden se concatenaba directamente en la consulta SQL: `oConn.Execute("UPDATE ... WHERE id=" & delId)`
- **Solución**: Se reemplazó por una consulta parametrizada con `ADODB.Command` y `CreateParameter`.

**b) Cambio de estado (`detail.asp` - `change_status`):**
- **Vulnerabilidad**: El nuevo estado se escapaba manualmente con `Replace(newStatus, "'", "''")` en lugar de usar parámetros, lo cual es propenso a bypass.
- **Solución**: Se implementó consulta parametrizada con `cmdCS.CommandText = "UPDATE cmms_work_orders SET status=? WHERE id=?"`

**c) Actualización de horas reales (`detail.asp` - `change_status` completado):**
- **Vulnerabilidad**: El `itemId` se concatenaba directamente en la subconsulta SUM.
- **Solución**: Se parametrizó completamente la actualización.

**d) Registro de tiempo (`detail.asp` - `add_time`):**
- **Vulnerabilidad**: El valor `tHours` se concatenaba directamente en la query: `VALUES (?, ?, " & tHours & ", ?, GETDATE())`
- **Solución**: Se agregó un parámetro tipado `@h` de tipo `adDouble (5)` para pasar las horas de forma segura.

#### 2. Módulo de Inventario (`modules/inventory/`)

**a) Acción de desactivación (`index.asp`):**
- **Vulnerabilidad**: El ID se concatenaba directamente en la consulta de baja lógica.
- **Solución**: Se implementó consulta parametrizada.

### Resumen de Correcciones

| Archivo | Línea | Vulnerabilidad | Severidad | Estado |
|---|---|---|---|---|
| `modules/work_orders/index.asp` | 17 | SQLi en `action=delete` | **Crítico** | Corregido |
| `modules/work_orders/detail.asp` | 59 | SQLi en `change_status` | **Crítico** | Corregido |
| `modules/work_orders/detail.asp` | 47 | SQLi en `add_time (tHours)` | **Crítico** | Corregido |
| `modules/work_orders/detail.asp` | 62 | SQLi en `completed_at` update | **Alto** | Corregido |
| `modules/work_orders/detail.asp` | 53 | SQLi en `actual_hours` update | **Alto** | Corregido |
| `modules/inventory/index.asp` | 17 | SQLi en `action=delete` | **Crítico** | Corregido |

### Pendiente para Próximas Auditorías
- Revisar filtros de búsqueda (`LIKE '%...%'`) en `index.asp` de work_orders e inventory que aún usan escape manual con `Replace`. Aunque funcional, se recomienda migrar a ADODB.Command para consistencia total.
- Agregar validación de token CSRF en las acciones POST de `detail.asp`.

## 2026-05-30: Sincronización Segura de Esquemas de Base de Datos y Auditoría de Módulos
- **Sanitización del Esquema SQL**: Se corrigió una anomalía de codificación de fin de archivo en `sql/mssql.sql` que provocaba caracteres espaciados corruptos que podían interferir con la correcta inicialización de la base de datos de manera segura.
- **Auditoría del Módulo de Solicitudes de Trabajo (`work_requests`)**:
  - Se verificó que la inserción de solicitudes (`form.asp`) esté protegida mediante tokens CSRF (`ValidateCSRF()`) para prevenir ataques de falsificación de peticiones en sitios cruzados.
  - Se validó el uso estricto de parámetros con `ADODB.Command` para evitar vectores de inyección SQL (SQLi) al registrar el título y la descripción del reporte de problemas.
  - Se auditó el acceso del archivo de aprobación (`approve.asp`) garantizando el chequeo de permisos administrativos/supervisores (`IsSupervisorOrAdmin()`) antes de permitir la resolución de cualquier solicitud y generación automática de la Orden de Trabajo.

## 2026-05-24: Mitigación de Inyección SQL y Manejo de Errores

### 1. Manejo de Errores Global (Debug Mode)
Se implementó un mecanismo en `core/functions.asp` para ocultar las excepciones de base de datos a los usuarios normales. En producción, el sistema muestra mensajes genéricos y detiene la ejecución amigablemente en caso de falla crítica, en lugar de exponer la estructura de la base de datos o el código de error VBScript. Los errores técnicos solo se mostrarán si el sistema detecta que el usuario es un administrador (basado en el rol) o si la variable de configuración de debug está activa.

### 2. Refactorización de Inyecciones SQL
Durante la auditoría inicial se encontraron los siguientes vectores de ataque potenciales en las eliminaciones lógicas (`action=delete`), donde el parámetro ID recibido por la URL (`Request.QueryString`) no estaba siendo parametrizado ni forzado de forma nativa en la capa de ejecución (aunque la función `QSInt` devuelve 0 si no es numérico, es una mejor práctica parametrizarlo explícitamente cuando es posible o asegurar su validación).
- Se ha actualizado la lógica de borrado lógico y cambios de estado en:
  - `modules/admin/settings.asp`
  - `modules/assets_module/index.asp`
  - `modules/inventory/index.asp`
  - `modules/inventory/movements.asp`
  - `modules/plants/index.asp`
  - `modules/users/index.asp`
  - `modules/work_orders/index.asp`
  - `modules/work_orders/detail.asp`
  - `modules/work_orders/form.asp`
  - `modules/users/form.asp`
  - `modules/users/profile.asp`

### 3. Configuración Segura
- Las credenciales de la base de datos en `config/database.asp` han sido protegidas utilizando las mejores prácticas de IIS. Se instruirá (en el README) cómo mover las contraseñas fuera del web root o utilizando el Administrador de Credenciales/Variables de Servidor si el administrador del servidor lo prefiere.

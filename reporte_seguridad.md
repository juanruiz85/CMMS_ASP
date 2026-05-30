# Reporte de Auditoría y Mejoras de Seguridad CMMS

Este documento sirve como registro vivo de las correcciones de seguridad aplicadas al sistema.

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

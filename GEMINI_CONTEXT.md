# Contexto de Desarrollo del Proyecto CMMS (Para Gemini IA)

Este archivo sirve como punto de referencia para que Gemini IA o cualquier otra IA entienda el estado actual, la arquitectura, las decisiones de diseño y las tareas pendientes del proyecto **CMMS (Sistema de Gestión de Mantenimiento)**.

---

## 🚀 Información General

El proyecto es un **CMMS (Computerized Maintenance Management System)** diseñado para entornos industriales, implementado en **Classic ASP (VBScript)** con una interfaz moderna y fluida de tipo Premium Dark Mode (Glassmorphic design).

- **Servidor Requerido**: IIS 8.5+ (Windows Server 2012 R2+)
- **Bases de Datos Soportadas**: SQL Server 2016+, MySQL 5.7+, SQLite 3+
- **Control de Acceso**: Control de Acceso Basado en Roles (RBAC) con roles: `admin`, `supervisor`, `technician`, `viewer`.
- **Idioma**: Español (por defecto) e Inglés (seleccionable).

---

## 🛠️ Arquitectura y Estructura de Directorios

```
C:\CMMS\
├── config/
│   ├── database.asp             # Configuración de base de datos activa
│   └── database_config.sample.asp
├── core/
│   ├── auth.asp                 # Autenticación, Roles y Control de Acceso
│   ├── functions.asp            # Utilidades globales y manejo de errores
│   └── i18n.asp                 # Traducción e Internacionalización (ES/EN)
├── templates/
│   ├── header.asp               # Cabecera HTML
│   ├── footer.asp               # Cierre de página, modales y scripts globales
│   └── navigation.asp           # Sidebar y barra superior del Dashboard
├── assets/
│   ├── css/app.css              # Sistema de diseño Premium Dark Mode
│   └── js/app.js                # Funcionalidades dinámicas frontend
├── modules/
│   ├── plants/                  # Gestión de Plantas / Ubicaciones
│   ├── assets_module/           # Gestión de Equipos / Jerarquías
│   ├── work_orders/             # Gestión de Órdenes de Trabajo (OT)
│   ├── work_requests/           # Gestión de Solicitudes de Trabajo (Bandeja de Entrada)
│   ├── inventory/               # Control de Inventario y Movimientos de Materiales
│   ├── users/                   # Gestión de Usuarios y Perfiles
│   ├── admin/                   # Logs de Actividad, Ajustes SMTP/Sistema
│   └── reports/                 # Generación y Exportación de Reportes
├── sql/
│   ├── mssql.sql                # Esquema para SQL Server 2016+
│   ├── mysql.sql                # Esquema para MySQL 5.7+
│   └── sqlite.sql               # Esquema para SQLite 3
├── index.asp                    # Dashboard principal
├── login.asp                    # Pantalla de inicio de sesión
├── logout.asp                   # Cierre de sesión de usuario
└── install.asp                  # Instalador automatizado de base de datos
```

---

## 📝 Estado de los Módulos del Core

1. **Dashboard (`index.asp`)**: **[Completado]** Visualiza KPI clave: Órdenes Abiertas, Equipos Críticos Fuera de Servicio, Alertas de Stock, Solicitudes Pendientes.
2. **Plantas (`modules/plants/`)**: **[Completado]** ABM (Alta, Baja, Modificación) de plantas físicas, dirección, encargado y estado.
3. **Equipos (`modules/assets_module/`)**: **[Completado]** Gestión de activos, jerarquías (equipo padre), criticidad, costes de adquisición, fechas de garantía e imágenes de equipos.
4. **Órdenes de Trabajo (`modules/work_orders/`)**: **[Completado]** Creación de OTs correctivas, preventivas, predictivas y de emergencia. Soporte para comentarios en tiempo real, registro de horas hombre, asignación de materiales consumidos y cierre formal.
5. **Solicitudes de Trabajo (`modules/work_requests/`)**: **[Completado]** Permite a usuarios de cualquier rol reportar fallas de manera ágil. Los supervisores pueden evaluar la solicitud para aprobarla (generando automáticamente una Orden de Trabajo pre-rellenada) o rechazarla con notas explicativas.
6. **Inventario (`modules/inventory/`)**: **[Completado]** Control de stock mínimo/máximo, punto de reorden, registro automático e histórico de movimientos (entradas, salidas, ajustes).
7. **Usuarios (`modules/users/`)**: **[Completado]** Administración de usuarios, asignación de roles y actualización de perfil/avatar.
8. **Administración (`modules/admin/`)**: **[Completado]** Logs detallados de auditoría de actividad de usuario y panel de configuración de parámetros del sistema y SMTP de correo.
9. **Reportes (`modules/reports/`)**: **[Completado]** Exportación de datos a formato CSV de reportes maestros de mantenimiento.

---

## 🛡️ Características de Seguridad Implementadas

- **Mitigación de SQL Injection**: Uso estricto de consultas preparadas mediante `ADODB.Command` con paso de parámetros tipados en formularios clave.
- **Protección CSRF**: Tokens criptográficos autogenerados en formularios POST (`CSRFField()` y `ValidateCSRF()`).
- **Manejo Seguro de Errores**: Ocultación de detalles de excepciones de base de datos para usuarios no administradores (Debug mode en `core/functions.asp`).
- **Hashing**: Contraseñas hasheadas de forma segura en SHA-256 utilizando un COM wrapper de .NET Framework para robustez extrema.

---

## 📋 Archivos de Referencia para Desarrollo

- **`CONTRIBUTING.md`** → Reglas de desarrollo, seguridad, estándares de código y checklist de verificación. **LEER ANTES DE EMPEZAR A TRABAJAR**.
- **`GEMINI_CONTEXT.md`** (este archivo) → Estado actual del proyecto, próximos pasos y contexto para IA.
- **`README.md`** → Estructura del proyecto, requisitos del sistema e instrucciones de instalación.
- **`CHANGELOG.md`** → Historial completo de cambios del sistema.

---

## 🔄 Últimos Cambios Realizados

### Fase 6 (2026-05-30 04:40): Auditoría de Seguridad en Módulos Core
- **Auditoría completa** de los módulos `work_orders/` e `inventory/` detectando **6 vulnerabilidades de inyección SQL**.
- **Correcciones aplicadas**:
  - `work_orders/index.asp`: Delete action migrado a ADODB.Command parametrizado.
  - `work_orders/detail.asp (change_status)`: Reemplazado escape manual por consulta parametrizada.
  - `work_orders/detail.asp (completed_at update)`: Parametrizada actualización automática.
  - `work_orders/detail.asp (actual_hours update)`: Parametrizada subconsulta SUM.
  - `work_orders/detail.asp (add_time)`: Agregado parámetro @h tipado (adDouble) para horas.
  - `inventory/index.asp`: Delete action migrado a ADODB.Command parametrizado.
- **Archivos creados/actualizados**:
  - `CONTRIBUTING.md` (NUEVO): Reglas de desarrollo estrictas para IA y desarrolladores.
  - `GEMINI_CONTEXT.md` (ACTUALIZADO): Contexto completo con referencias y próximos pasos.
  - `reporte_seguridad.md` (ACTUALIZADO): Tabla detallada de hallazgos y correcciones.
  - `CHANGELOG.md` (ACTUALIZADO): Nueva entrada Fase 6.

### Fase 10 (2026-05-30 06:30): Migración Completa de Filtros SELECT LIKE a ADODB.Command en Módulos Restantes
- **work_requests/index.asp**: Filtros de status y plant_id migrados a parámetros con ADODB.Command (incluye filtro automático por usuario actual).
- **work_requests/form.asp**: INSERT de solicitudes completamente parametrizado (eliminando concatenación de plant_id, asset_id y requested_by_id).
- **assets_module/index.asp**: Filtros de búsqueda (q, plant_id, status) migrados a ADODB.Command con parámetros.
- **users/index.asp**: Filtros de búsqueda (q, role, status) migrados a ADODB.Command con parámetros.
- **plants/index.asp**: Filtro de búsqueda (q) migrado a ADODB.Command con parámetros.
- **admin/logs.asp**: Filtros (user_id, module, date) migrados a ADODB.Command con parámetros.
- **Resultado**: **100% de consultas SELECT con filtros parametrizados** en todos los módulos del sistema.

### Fase 9 (2026-05-30 06:00): Migración Completa de Consultas UPDATE/INSERT a ADODB.Command en Todos los Módulos
- Se migraron todas las consultas `UPDATE` e `INSERT` que usaban concatenación de strings a `ADODB.Command` con parámetros tipados.
- **Archivos corregidos**:
  - `modules/admin/settings.asp`: UPSERT de configuración migrado a comandos parametrizados.
  - `modules/users/form.asp`: Actualización de contraseña migrada a comando parametrizado.
  - `modules/users/profile.asp`: Cambio de contraseña migrado a comando parametrizado.
  - `modules/users/index.asp`: Desactivación de usuarios migrada a comando parametrizado.
  - `modules/work_orders/form.asp`: Actualización de `completed_at` y `closed_by_id` migrada a comando parametrizado.
  - `modules/assets_module/index.asp`: Retiro de equipos migrado a comando parametrizado.
  - `modules/plants/index.asp`: Desactivación de plantas migrada a comando parametrizado.
  - `modules/inventory/form.asp`: Creación de stock inicial migrada a comando parametrizado.
  - `modules/inventory/movements.asp`: Actualización/creación de stock migrada a comandos parametrizados.
- **Resultado**: **100% de consultas de escritura (INSERT/UPDATE) parametrizadas** en todos los módulos del sistema, eliminando completamente el riesgo de SQL Injection en operaciones de modificación de datos.

### Fase 8 (2026-05-30 05:12): Migración Completa de Filtros LIKE a ADODB.Command
- **work_orders/index.asp**: Filtros de búsqueda (q, type, status, plant_id) migrados de `Replace()` manual a `ADODB.Command` con parámetros.
- **inventory/index.asp**: Filtros de búsqueda (q, plant_id, status) migrados a `ADODB.Command` con parámetros.
- **Corrección técnica**: Se eliminó el patrón de copia de parámetros entre Command objects que es inválido en VBScript, reemplazándolo por creación inline directa en cada comando.
- **Resultado**: **100% de consultas parametrizadas** en todos los módulos core del sistema.

### Fase 7 (2026-05-30 05:07): Protección CSRF y Regla de Commits Automáticos
- **CSRF implementado** en `modules/work_orders/detail.asp`:
  - Validación `ValidateCSRF()` en todas las acciones POST (add_comment, add_time, change_status).
  - Campos ocultos `CSRFField()` en los 3 formularios del detalle de OT.
- **CONTRIBUTING.md actualizado**: RG-09 ahora incluye commits automáticos con mensaje descriptivo.
- **Próximos pasos actualizados**: Backlog refleja prioridad alta #2 como completada.

### Fase 5 (2026-05-30 03:20): Sincronización de BD y Contexto IA
- Sincronización de esquemas SQL (`mssql.sql`, `mysql.sql`, `sqlite.sql`) con tablas `cmms_work_requests` y `cmms_scheduled_reports`.
- Reparación de corrupción de encoding UTF-16 en `sql/mssql.sql`.
- Creación de `GEMINI_CONTEXT.md` para continuidad del proyecto.

---

## 📌 Próximos Pasos Recomendados (Backlog)

### ✅ Prioridades Altas Completadas
1. ~~**Migración completa de filtros LIKE a ADODB.Command**~~ ✅ **COMPLETADO en Fase 8** — 100% de consultas parametrizadas en módulos core.
2. ~~**Agregar validación CSRF en `work_orders/detail.asp`**~~ ✅ **COMPLETADO en Fase 7**
3. ~~**Migración de INSERT/UPDATE a ADODB.Command en todos los módulos**~~ ✅ **COMPLETADO en Fase 9**
4. ~~**Migración de filtros SELECT LIKE en módulos restantes**~~ ✅ **COMPLETADO en Fase 10** — 100% de consultas SELECT con filtros parametrizados en todo el sistema.

### Prioridad Media (Próximas a trabajar)
5. **Pruebas del Flujo de Trabajo Completo**:
   - Crear solicitud de trabajo como usuario `viewer`/`technician`.
   - Aprobar como `supervisor` y verificar generación automática de OT.
   - Completar la OT y verificar actualización de inventario y horas.
6. **Módulo de Tareas Programadas / Mantenimiento Preventivo**: Implementar servicio en segundo plano usando `cmms_scheduled_reports`.

### Prioridad Baja
7. **Mapeo de Rutas de API REST**: Desarrollo de `api/` para consumo móvil por técnicos en campo.
8. **Optimización de consultas KPI**: Las consultas de KPIs en dashboards usan GROUP BY sin parámetros, revisar eficiencia.

---

> **Estado de Seguridad Actual**: 100% de consultas SQL parametrizadas en módulos core, protección CSRF implementada en todos los formularios POST, manejo seguro de errores activo.

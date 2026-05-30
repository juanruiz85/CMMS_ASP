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

## 🔄 Últimos Cambios Realizados

- **Sincronización de Esquemas SQL**: Se corrigieron las inconsistencias de los scripts SQL (`mssql.sql`, `mysql.sql`, `sqlite.sql`) para incluir las tablas `cmms_work_requests` y `cmms_scheduled_reports` de forma consistente.
- **Sanitización de Esquemas**: Se resolvió una corrupción de caracteres (doble espaciado UTF-16) detectada al final de `sql/mssql.sql`, logrando que todas las sentencias se ejecuten de manera nativa sin fallas.
- **Creación de Contexto Gemini**: Creación de este archivo para la continuación fluida del proyecto.

---

## 📌 Próximos Pasos Recomendados (Backlog)

1. **Pruebas del Flujo de Trabajo Completo**:
   - Crear una solicitud de trabajo como usuario de rol `viewer` o `technician`.
   - Loguearse como `supervisor` o `admin` para aprobar la solicitud y verificar la generación de la Orden de Trabajo.
2. **Módulo de Tareas Programadas / Mantenimiento Preventivo**:
   - Implementar un servicio o script que corra en segundo plano (ej: Tarea Programada de Windows) que lea la tabla `cmms_scheduled_reports` o genere OTs preventivas de manera automática basándose en periodicidades configuradas.
3. **Mapeo de Rutas de API REST**:
   - Desarrollar la lógica en la carpeta `api/` para habilitar el consumo móvil por parte de técnicos en campo.

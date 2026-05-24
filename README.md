# CMMS — Sistema de Gestión de Mantenimiento

Sistema CMMS (Computerized Maintenance Management System) desarrollado en **Classic ASP (VBScript)** con interfaz moderna dark mode premium. Compatible con **IIS 8.5+ (Windows Server 2012 R2+)** y **SQL Server 2016+**.

---

## Requisitos del Sistema

| Componente | Versión Mínima |
|---|---|
| Windows Server | 2012 R2 |
| IIS | 8.5 |
| SQL Server | 2016 |
| SQL Server Native Client | SQLNCLI11 |
| .NET Framework | 4.5 (para hashing SHA-256) |
| Navegador | Chrome 80+ / Firefox 75+ / Edge 80+ |

---

## Instalación en IIS

### 1. Configurar IIS

1. Abrir **IIS Manager**
2. Crear un nuevo **Application Pool**:
   - Nombre: `CMMS`
   - Versión .NET: **Sin código administrado** (Classic ASP no usa .NET)
   - Modo de canalización: **Integrado**
3. Crear un nuevo **Sitio Web** o **Aplicación Virtual**:
   - Alias: `CMMS`
   - Ruta física: `C:\CMMS`
   - Application Pool: `CMMS`
4. Habilitar **ASP** en IIS:
   - Roles del Servidor → Servidor Web → Características HTTP → ASP ✓
5. Habilitar errores de script detallados (para depuración):
   - Sitio → ASP → Comportamiento de depuración → Enviar errores al explorador = True

### 2. Configurar permisos de carpeta

```powershell
# Dar permisos al usuario IIS_IUSRS sobre la carpeta CMMS
icacls "C:\CMMS" /grant "IIS_IUSRS:(OI)(CI)RX" /T
icacls "C:\CMMS\uploads" /grant "IIS_IUSRS:(OI)(CI)M" /T
icacls "C:\CMMS\config" /grant "IIS_IUSRS:(OI)(CI)M" /T
```

### 3. Ejecutar el instalador

1. Navegar a: `http://tu-servidor/CMMS/install.asp`
2. Seleccionar: **SQL Server**
3. Ingresar datos de conexión:
   - Servidor: nombre del servidor SQL (ej. `SQLSERVER01` o `.\SQLEXPRESS`)
   - Base de datos: nombre de la BD a crear (ej. `CMMS_DB`)
   - Usuario: usuario SQL con permisos de creación
   - Contraseña: contraseña del usuario SQL
4. Hacer clic en **Instalar**
5. El sistema creará todas las tablas y el usuario administrador

### 4. Credenciales por defecto

```
Usuario: admin
Contraseña: Admin@123!
```

> ⚠️ **IMPORTANTE**: Cambie la contraseña del administrador inmediatamente después del primer acceso.

---

## Estructura del Proyecto

```
C:\CMMS\
├── sql/                    # Scripts SQL por motor de BD
│   ├── mssql.sql          # SQL Server schema
│   ├── mysql.sql          # MySQL schema
│   └── sqlite.sql         # SQLite schema
├── config/
│   ├── database.asp       # Configuración activa (generado por instalador)
│   └── database_config.sample.asp
├── core/
│   ├── auth.asp           # Autenticación y RBAC
│   ├── functions.asp      # Funciones helper globales
│   ├── i18n.asp           # Internacionalización (ES/EN)
│   └── installer.asp      # Lógica del instalador
├── templates/
│   ├── header.asp         # Cabecera HTML
│   ├── footer.asp         # Pie de página y scripts
│   ├── navigation.asp     # Sidebar + topbar
│   └── default-dashboard.asp
├── assets/
│   ├── css/app.css        # Sistema de diseño dark mode
│   ├── js/app.js          # JavaScript global
│   └── images/
├── modules/
│   ├── plants/            # Gestión de plantas
│   ├── assets_module/     # Gestión de equipos
│   ├── work_orders/       # Órdenes de trabajo
│   ├── inventory/         # Inventario y materiales
│   ├── users/             # Gestión de usuarios
│   ├── admin/             # Panel de administración
│   └── reports/           # Reportes y analytics
├── api/                   # API REST
├── uploads/               # Archivos subidos
├── index.asp              # Dashboard principal
├── login.asp              # Inicio de sesión
├── logout.asp             # Cierre de sesión
├── install.asp            # Instalador web
├── CHANGELOG.md           # Historial de cambios
└── README.md              # Este archivo
```

---

## Roles del Sistema

| Rol | Descripción | Acceso |
|---|---|---|
| `admin` | Administrador del sistema | Total |
| `supervisor` | Supervisor de mantenimiento | Módulos + reportes |
| `technician` | Técnico de mantenimiento | OTs asignadas + inventario |
| `viewer` | Solo lectura | Consultas |

---

## API REST

### Autenticación

```
POST /CMMS/api/auth.asp
Body: {"username": "admin", "password": "Admin@123!"}
Response: {"success": true, "token": "API_KEY_HERE"}
```

### Endpoints principales

```
GET  /CMMS/api/assets.asp        — Listar equipos
POST /CMMS/api/work_orders.asp   — Crear orden de trabajo
GET  /CMMS/api/inventory.asp     — Consultar inventario
GET  /CMMS/api/reports.asp?id=1  — Ejecutar reporte
```

---

## Idiomas Soportados

- 🇪🇸 **Español** (por defecto)
- 🇺🇸 **English** (disponible en Perfil de Usuario → Preferencias)

---

## Seguridad

- Contraseñas hasheadas con **SHA-256** (.NET Framework COM)
- Queries parametrizadas con **ADODB.Command** (protección SQL Injection)
- Validación de entradas en cliente y servidor
- Sistema RBAC (Control de Acceso Basado en Roles)
- Logs de auditoría completos
- Tokens CSRF en formularios POST
- Sesiones con timeout configurable

---

## Soporte

Sistema de código abierto. Para reportar problemas, revisar el CHANGELOG.md para el historial de cambios.

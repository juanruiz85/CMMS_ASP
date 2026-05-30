# Changelog del Sistema CMMS

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

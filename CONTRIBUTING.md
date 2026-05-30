# Reglas de Desarrollo para el Proyecto CMMS

Este documento establece las reglas y pautas estrictas para todo desarrollo, modificación o auditoría del sistema CMMS. **Cada vez que una IA o desarrollador realice cambios, debe seguir estas reglas rigurosamente.**

---

## 📋 Reglas Generales

### RG-01: Documentación
- **Siempre** actualizar `README.md`, `CHANGELOG.md` y `GEMINI_CONTEXT.md` después de cada bloque de cambios.
- Documentar en el código: toda función debe tener un comentario describiendo qué hace, qué parámetros recibe y qué retorna.
- Usar comentarios en inglés para el código y español para la interfaz de usuario.

### RG-02: Estructura del Proyecto
- Respetar la estructura de directorios definida en `README.md`.
- No crear archivos en la raíz del proyecto a menos que sea estrictamente necesario.
- Los módulos van dentro de `modules/` con su propia carpeta (ej: `modules/work_orders/`).

### RG-03: Idioma del Sistema
- La interfaz de usuario debe estar en **español** por defecto.
- Todo el sistema debe soportar **cambio a inglés** mediante el sistema de internacionalización en `core/i18n.asp`.
- Usar la función `T("key")` para todas las cadenas de texto visibles al usuario.
- Agregar las traducciones al archivo `core/i18n.asp`.

### RG-04: Diseño de Interfaz
- **Siempre** revisar varias páginas existentes antes de crear una nueva para mantener consistencia visual.
- El sistema de diseño es **Dark Mode Premium** con glassmorphism.
- No alterar el diseño global definido en `assets/css/app.css`.
- Usar los componentes y clases existentes (`.card`, `.btn`, `.form-group`, `.badge`, etc.).
- No usar CSS inline a menos que sea estrictamente necesario para un caso específico.

### RG-05: Seguridad (Prioridad Máxima)
- **TODAS** las consultas a la base de datos (INSERT, UPDATE, SELECT, DELETE) deben usar `ADODB.Command` con parámetros tipados.
- **NUNCA** concatenar valores directamente en cadenas SQL.
- **SIEMPRE** validar tokens CSRF en formularios POST usando `ValidateCSRF()` y `CSRFField()`.
- **SIEMPRE** sanitizar salidas HTML con `HtmlEncode()`.
- Nunca exponer errores de base de datos al usuario final (usar `CheckError()`).
- Las contraseñas deben hashearse con `SHA256Hash()` usando salt (ver `core/functions.asp`).
- Toda acción de escritura debe ser registrada con `LogActivity()`.

### RG-06: Compatibilidad Multi-Base de Datos
- Toda consulta SQL debe ser compatible con **SQL Server**, **MySQL** y **SQLite**.
- NO usar funciones específicas de un motor sin verificar compatibilidad:
  - ❌ `GETDATE()` → usar en SQL Server pero no en MySQL (usar `NOW()` o `CURRENT_TIMESTAMP`)
  - ❌ `IDENT_CURRENT()` → no usar, es específico de SQL Server
  - ❌ `OFFSET...FETCH NEXT` → solo SQL Server 2012+
  - ✅ Usar parámetros `?` con ADODB.Command (el driver adapta automáticamente)
- Al crear esquemas, replicar los cambios en los 3 archivos: `sql/mssql.sql`, `sql/mysql.sql`, `sql/sqlite.sql`.

### RG-07: Control de Acceso (RBAC)
- Toda página o acción debe verificar autenticación con `CheckAuth()`.
- Verificar permisos según el rol usando:
  - `IsAdmin()` - Solo administradores
  - `IsSupervisorOrAdmin()` - Supervisores y administradores
  - Cualquier usuario autenticado puede acceder a funcionalidades básicas
- Roles definidos: `admin`, `supervisor`, `technician`, `viewer`.

### RG-08: Estilo de Código ASP
- Usar sangría consistente (2 espacios).
- Declarar todas las variables con `Dim`.
- Usar comentarios `' ---` para secciones importantes.
- Mantener las líneas de código legibles (máximo 120 caracteres por línea).
- Las funciones deben seguir el patrón: nombre en PascalCase.
- Las variables en camelCase (ej: `itemId`, `filterQuery`).

### RG-09: Gestión de Cambios y Commits Automáticos
- **Siempre hacer commit con mensaje descriptivo automático** después de completar un bloque de cambios.
- El mensaje de commit debe incluir:
  - Fase/número de iteración
  - Qué se hizo (archivos modificados/creados)
  - Por qué se hizo (motivo del cambio)
- Usar el formato: `Fase X: Breve descripción del cambio`
- Ejemplo: `Fase 7: Agregar validación CSRF en work_orders/detail.asp`
- Después de una sesión de desarrollo:
  1. Actualizar `CHANGELOG.md` con una entrada descriptiva.
  2. Actualizar `README.md` si cambió la estructura o requisitos.
  3. Actualizar `GEMINI_CONTEXT.md` con el estado actual y próximos pasos.
  4. Actualizar `CONTRIBUTING.md` si se identificaron nuevas reglas.
  5. Hacer commit con mensaje automático descriptivo.

### RG-10: Conexiones a Base de Datos
- Siempre cerrar conexiones con `CloseConnection(oConn)`.
- Cerrar RecordSets después de usarlos.
- Usar `On Error Resume Next` / `On Error GoTo 0` para manejo de errores controlado.

---

## 🚨 Checklist de Verificación Pre-Entrega

Antes de finalizar cualquier bloque de cambios, verificar:

- [ ] ¿Todas las consultas SQL usan ADODB.Command con parámetros?
- [ ] ¿Todos los formularios POST tienen token CSRF?
- [ ] ¿Todas las salidas HTML están sanitizadas con `HtmlEncode()`?
- [ ] ¿Los permisos de acceso están validados con `CheckAuth()`?
- [ ] ¿Los esquemas SQL están sincronizados en los 3 motores?
- [ ] ¿Las traducciones están agregadas en `core/i18n.asp`?
- [ ] ¿El código está comentado adecuadamente?
- [ ] ¿Se actualizó `CHANGELOG.md`?
- [ ] ¿Se actualizó `GEMINI_CONTEXT.md`?
- [ ] ¿Si aplica, se actualizó `README.md`?

---

## 🔄 Flujo de Trabajo Recomendado para IA

1. **Leer este archivo** (`CONTRIBUTING.md`) primero para conocer las reglas.
2. **Leer `GEMINI_CONTEXT.md`** para entender el estado actual del proyecto.
3. **Leer `README.md`** para conocer la estructura y requisitos.
4. **Revisar páginas existentes** para mantener consistencia de diseño.
5. **Implementar cambios** siguiendo todas las reglas de seguridad y estilo.
6. **Verificar checklist** de pre-entrega.
7. **Actualizar documentación** (`CHANGELOG.md`, `GEMINI_CONTEXT.md`, `README.md`).

---

## 🎯 Próximas Prioridades (de GEMINI_CONTEXT.md)

1. **Pruebas del Flujo de Trabajo Completo**: Crear solicitud → Aprobar → Generar OT
2. **Módulo de Tareas Programadas / Mantenimiento Preventivo**: Servicio en segundo plano usando `cmms_scheduled_reports`
3. **Mapeo de Rutas de API REST**: Desarrollo de `api/` para consumo móvil
4. **Migración completa de filtros LIKE a ADODB.Command**: En `work_orders/index.asp` e `inventory/index.asp`
5. **Agregar validación CSRF**: En acciones POST de `work_orders/detail.asp`
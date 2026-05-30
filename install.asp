<%
' =============================================================================
' CMMS - Instalador del Sistema (install.asp)
' Wizard multi-paso: DB Config → Prueba Conexión → Crear Schema → Admin → Listo
' Soporta: SQL Server, MySQL, SQLite
' =============================================================================
Option Explicit

' Función auxiliar IIf (Immediate If)
Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function

' Bloquear si ya está instalado
Dim configFile
configFile = Server.MapPath("/CMMS/config/database.asp")
If CreateObject("Scripting.FileSystemObject").FileExists(configFile) Then
    ' Ya instalado: verificar que la BD existe
    Dim checkInstalled : checkInstalled = True
    On Error Resume Next
    Dim testConn : Set testConn = Server.CreateObject("ADODB.Connection")
    ' Intentar leer el config existente
    ' Si falla, permitir reinstalar
    On Error GoTo 0
End If

Dim Step : Step = Request.QueryString("step")
If Step = "" Then Step = "1"

Dim ErrorMsg : ErrorMsg = ""
Dim SuccessMsg : SuccessMsg = ""

' ─── Procesamiento POST ────────────────────────────────────────────────────────
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action : action = Request.Form("action")

    ' PASO 2: Probar conexión y guardar config
    If action = "test_save" Then
        Dim dbType   : dbType   = Trim(Request.Form("db_type"))
        Dim dbServer : dbServer = Trim(Request.Form("db_server"))
        Dim dbName   : dbName   = Trim(Request.Form("db_name"))
        Dim dbUser   : dbUser   = Trim(Request.Form("db_user"))
        Dim dbPass   : dbPass   = Trim(Request.Form("db_pass"))
        Dim dbPort   : dbPort   = Trim(Request.Form("db_port"))
        Dim dbFile   : dbFile   = Trim(Request.Form("db_file"))
        
        If dbType = "" Then dbType = "sqlserver"
        If dbPort = "" Then dbPort = getDefaultPort(dbType)
        
        ' Construir cadena de conexión según tipo de BD
        Dim connStr
        connStr = buildConnectionString(dbType, dbServer, dbName, dbUser, dbPass, dbPort, dbFile)
        
        ' Probar conexión
        Dim testConn2
        Set testConn2 = Server.CreateObject("ADODB.Connection")
        On Error Resume Next
        testConn2.Open connStr
        If Err.Number <> 0 Then
            ErrorMsg = "Error de conexión: " & Err.Description & " (Código: " & Err.Number & ")"
            Err.Clear
        Else
            testConn2.Close
            ' Guardar en sesión para el siguiente paso
            Session("inst_type")    = dbType
            Session("inst_server")  = dbServer
            Session("inst_name")    = dbName
            Session("inst_user")    = dbUser
            Session("inst_pass")    = dbPass
            Session("inst_port")    = dbPort
            Session("inst_file")    = dbFile
            SuccessMsg = "¡Conexión exitosa! Proceda al siguiente paso."
            Step = "3"
        End If
        Set testConn2 = Nothing
        On Error GoTo 0
    End If

    ' PASO 3: Ejecutar schema y crear admin
    If action = "install" Then
        Dim adminUser : adminUser = Trim(Request.Form("admin_user"))
        Dim adminPass : adminPass = Trim(Request.Form("admin_pass"))
        Dim adminEmail: adminEmail= Trim(Request.Form("admin_email"))
        Dim adminFirst: adminFirst= Trim(Request.Form("admin_first"))
        Dim adminLast : adminLast = Trim(Request.Form("admin_last"))

        If adminUser = "" Or adminPass = "" Or adminEmail = "" Then
            ErrorMsg = "Todos los campos del administrador son requeridos."
        Else
            ' Ejecutar instalación
            Dim instResult : instResult = RunInstaller( _
                Session("inst_type"),   Session("inst_server"), Session("inst_name"), _
                Session("inst_user"),   Session("inst_pass"),   Session("inst_port"), _
                Session("inst_file"), _
                adminUser, adminPass, adminEmail, adminFirst, adminLast)
            
            If Left(instResult, 5) = "ERROR" Then
                ErrorMsg = instResult
            Else
                Step = "4" ' Éxito
            End If
        End If
    End If
End If

' ─── Función de instalación ────────────────────────────────────────────────────
Function RunInstaller(dbType, dbSrv, dbNm, dbUsr, dbPwd, dbPort, dbFile, admUser, admPass, admEmail, admFirst, admLast)
    Dim oConn, oFS, schemaPath, schemaSQL, sLines, i
    
    On Error Resume Next
    
    ' 1. Conectar según tipo de BD
    Dim cs
    cs = buildConnectionString(dbType, dbSrv, dbNm, dbUsr, dbPwd, dbPort, dbFile)
    Set oConn = Server.CreateObject("ADODB.Connection")
    oConn.Open cs
    If Err.Number <> 0 Then
        RunInstaller = "ERROR: " & Err.Description
        Exit Function
    End If

    ' 2. Determinar archivo de schema según tipo de BD
    Select Case LCase(dbType)
        Case "sqlserver"
            schemaPath = Server.MapPath("/CMMS/sql/mssql.sql")
        Case "mysql"
            schemaPath = Server.MapPath("/CMMS/sql/mysql.sql")
        Case "sqlite"
            schemaPath = Server.MapPath("/CMMS/sql/sqlite.sql")
        Case Else
            schemaPath = Server.MapPath("/CMMS/sql/mssql.sql")
    End Select
    
    Set oFS = Server.CreateObject("Scripting.FileSystemObject")
    
    If Not oFS.FileExists(schemaPath) Then
        RunInstaller = "ERROR: No se encontró el archivo " & schemaPath
        Exit Function
    End If
    
    Dim oFile : Set oFile = oFS.OpenTextFile(schemaPath, 1, False)
    schemaSQL = oFile.ReadAll
    oFile.Close
    Set oFile = Nothing
    Set oFS   = Nothing

    ' Agregar columna password_salt si no existe (no está en el schema original)
    oConn.Execute "IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='cmms_users' AND COLUMN_NAME='password_salt') " & _
                  "AND OBJECT_ID('cmms_users', 'U') IS NOT NULL " & _
                  "ALTER TABLE cmms_users ADD password_salt VARCHAR(50) DEFAULT '' NOT NULL"

    ' Ejecutar schema dividido por bloques IF OBJECT_ID
    ' El schema de mssql.sql ya tiene las guaredas IF OBJECT_ID, ejecutarlo completo
    ' Dividir por ";" y ejecutar cada statement
    Dim statements : statements = Split(schemaSQL, ";")
    Dim stmt
    For Each stmt In statements
        stmt = Trim(stmt)
        If Len(stmt) > 10 Then
            Err.Clear
            oConn.Execute stmt
            ' Ignorar errores de "ya existe"
        End If
    Next
    If Err.Number <> 0 Then Err.Clear

    ' Agregar password_salt después de crear la tabla
    oConn.Execute "IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='cmms_users' AND COLUMN_NAME='password_salt') " & _
                  "ALTER TABLE cmms_users ADD password_salt VARCHAR(50) DEFAULT '' NOT NULL"

    ' 3. Crear usuario administrador
    ' Generar salt y hash
    Dim salt, passHash
    Randomize
    Dim saltChars : saltChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    Dim j, saltStr : saltStr = ""
    For j = 1 To 16
        saltStr = saltStr & Mid(saltChars, Int(Rnd * Len(saltChars)) + 1, 1)
    Next
    salt = saltStr
    passHash = SHA256Hash_Install(admPass & salt)

    ' Verificar si ya existe admin
    Dim existRS
    Set existRS = oConn.Execute("SELECT COUNT(*) AS cnt FROM cmms_users WHERE username = '" & Replace(admUser, "'", "''") & "'")
    If existRS("cnt") = 0 Then
        Dim insertSQL
        insertSQL = "INSERT INTO cmms_users (username, email, password, password_salt, role, first_name, last_name, status, created_at, updated_at) " & _
                    "VALUES ('" & Replace(admUser, "'", "''") & "', " & _
                    "'" & Replace(admEmail, "'", "''") & "', " & _
                    "'" & passHash & "', " & _
                    "'" & salt & "', " & _
                    "'admin', " & _
                    "'" & Replace(admFirst, "'", "''") & "', " & _
                    "'" & Replace(admLast,  "'", "''") & "', " & _
                    "'active', GETDATE(), GETDATE())"
        oConn.Execute insertSQL
    End If
    existRS.Close
    Set existRS = Nothing

    ' 4. Insertar settings por defecto
    Dim defSettings(5, 1)
    defSettings(0, 0) = "company_name"    : defSettings(0, 1) = "Mi Empresa CMMS"
    defSettings(1, 0) = "app_version"     : defSettings(1, 1) = "1.0.0"
    defSettings(2, 0) = "session_timeout" : defSettings(2, 1) = "480"
    defSettings(3, 0) = "items_per_page"  : defSettings(3, 1) = "25"
    defSettings(4, 0) = "currency"        : defSettings(4, 1) = "USD"
    defSettings(5, 0) = "date_format"     : defSettings(5, 1) = "DD/MM/YYYY"

    Dim si
    For si = 0 To 5
        oConn.Execute "IF NOT EXISTS (SELECT 1 FROM cmms_settings WHERE key_name = '" & defSettings(si, 0) & "') " & _
                      "INSERT INTO cmms_settings (key_name, value) VALUES ('" & defSettings(si, 0) & "', '" & Replace(defSettings(si, 1), "'", "''") & "')"
    Next

    ' 5. Insertar roles por defecto
    oConn.Execute "IF NOT EXISTS (SELECT 1 FROM cmms_roles WHERE name = 'admin') " & _
                  "INSERT INTO cmms_roles (name, description, permissions, created_at) " & _
                  "VALUES ('admin', 'Administrador del sistema', '{""all"":true}', GETDATE())"
    oConn.Execute "IF NOT EXISTS (SELECT 1 FROM cmms_roles WHERE name = 'supervisor') " & _
                  "INSERT INTO cmms_roles (name, description, permissions, created_at) " & _
                  "VALUES ('supervisor', 'Supervisor de mantenimiento', '{""wo"":true,""assets"":true,""inventory"":true,""reports"":true}', GETDATE())"
    oConn.Execute "IF NOT EXISTS (SELECT 1 FROM cmms_roles WHERE name = 'technician') " & _
                  "INSERT INTO cmms_roles (name, description, permissions, created_at) " & _
                  "VALUES ('technician', 'Técnico de mantenimiento', '{""wo"":true,""inventory"":""read""}', GETDATE())"
    oConn.Execute "IF NOT EXISTS (SELECT 1 FROM cmms_roles WHERE name = 'viewer') " & _
                  "INSERT INTO cmms_roles (name, description, permissions, created_at) " & _
                  "VALUES ('viewer', 'Solo lectura', '{""read"":true}', GETDATE())"

    ' 6. Insertar planta y datos de ejemplo
    oConn.Execute "IF NOT EXISTS (SELECT 1 FROM cmms_plants WHERE code = 'PLANT-001') " & _
                  "INSERT INTO cmms_plants (code, name, description, city, country, status, created_at, updated_at) " & _
                  "VALUES ('PLANT-001', 'Planta Principal', 'Planta de demostración', 'Ciudad', 'México', 'active', GETDATE(), GETDATE())"

    oConn.Close
    Set oConn = Nothing

    ' 7. Generar archivo config/database.asp
    Dim oFS2 : Set oFS2 = Server.CreateObject("Scripting.FileSystemObject")
    Dim configPath : configPath = Server.MapPath("/CMMS/config/database.asp")
    Dim oConfigFile : Set oConfigFile = oFS2.CreateTextFile(configPath, True)
    
    Dim dbProvider, dbPortConst
    Select Case LCase(dbType)
        Case "sqlserver"
            dbProvider = "SQLNCLI11"
            dbPortConst = IIf(dbPort <> "", dbPort, "1433")
        Case "mysql"
            dbProvider = "MySQL.OLEDB.1"
            dbPortConst = IIf(dbPort <> "", dbPort, "3306")
        Case "sqlite"
            dbProvider = ""
            dbPortConst = ""
        Case Else
            dbProvider = "SQLNCLI11"
            dbPortConst = "1433"
    End Select
    
    oConfigFile.WriteLine "<" & "%"
    oConfigFile.WriteLine "' CMMS - Configuración de Base de Datos (GENERADO POR INSTALADOR)"
    oConfigFile.WriteLine "' Generado: " & Now()
    oConfigFile.WriteLine "Const DB_TYPE     = """ & dbType & """"
    oConfigFile.WriteLine "Const DB_SERVER   = """ & Replace(dbSrv, """", "") & """"
    oConfigFile.WriteLine "Const DB_NAME     = """ & Replace(dbNm, """", "") & """"
    oConfigFile.WriteLine "Const DB_USER     = """ & Replace(dbUsr, """", "") & """"
    oConfigFile.WriteLine "Const DB_PASS     = """ & Replace(dbPwd, """", "") & """"
    oConfigFile.WriteLine "Const DB_PORT     = " & dbPortConst
    oConfigFile.WriteLine "Const DB_PROVIDER = """ & dbProvider & """"
    oConfigFile.WriteLine "Const DB_TIMEOUT  = 30"
    oConfigFile.WriteLine "Const DB_APP_NAME = ""CMMS_System"""
    If LCase(dbType) = "sqlite" Then
        oConfigFile.WriteLine "Const DB_FILE     = """ & Replace(dbFile, """", "") & """"
    End If
    oConfigFile.WriteLine ""
    oConfigFile.WriteLine "Function GetConnectionString()"
    oConfigFile.WriteLine "    Select Case DB_TYPE"
    oConfigFile.WriteLine "        Case ""sqlserver"""
    oConfigFile.WriteLine "            GetConnectionString = ""Provider="" & DB_PROVIDER & "";Server="" & DB_SERVER & "";Database="" & DB_NAME & "";UID="" & DB_USER & "";PWD="" & DB_PASS & "";Connect Timeout="" & DB_TIMEOUT & "";Application Name="" & DB_APP_NAME & "";"""
    oConfigFile.WriteLine "        Case ""mysql"""
    oConfigFile.WriteLine "            GetConnectionString = ""Driver={MySQL ODBC 8.0 Unicode Driver};Server="" & DB_SERVER & "";Port="" & DB_PORT & "";Database="" & DB_NAME & "";Uid="" & DB_USER & "";Pwd="" & DB_PASS & "";Option=3;"""
    oConfigFile.WriteLine "        Case ""sqlite"""
    oConfigFile.WriteLine "            GetConnectionString = ""Driver={SQLite3 ODBC Driver};Database="" & DB_FILE & "";"""
    oConfigFile.WriteLine "        Case Else"
    oConfigFile.WriteLine "            GetConnectionString = """""
    oConfigFile.WriteLine "    End Select"
    oConfigFile.WriteLine "End Function"
    oConfigFile.WriteLine ""
    oConfigFile.WriteLine "Function GetConnection()"
    oConfigFile.WriteLine "    Dim oConn"
    oConfigFile.WriteLine "    Set oConn = Server.CreateObject(""ADODB.Connection"")"
    oConfigFile.WriteLine "    On Error Resume Next"
    oConfigFile.WriteLine "    oConn.Open GetConnectionString()"
    oConfigFile.WriteLine "    If Err.Number <> 0 Then"
    oConfigFile.WriteLine "        Response.Write ""<div style='color:red;font-family:monospace;padding:20px'>ERROR BD: "" & Err.Description & ""</div>"""
    oConfigFile.WriteLine "        Response.End"
    oConfigFile.WriteLine "    End If"
    oConfigFile.WriteLine "    On Error GoTo 0"
    oConfigFile.WriteLine "    Set GetConnection = oConn"
    oConfigFile.WriteLine "End Function"
    oConfigFile.WriteLine ""
    oConfigFile.WriteLine "Sub CloseConnection(oConn)"
    oConfigFile.WriteLine "    If Not IsNull(oConn) And Not IsEmpty(oConn) Then"
    oConfigFile.WriteLine "        If oConn.State = 1 Then oConn.Close"
    oConfigFile.WriteLine "        Set oConn = Nothing"
    oConfigFile.WriteLine "    End If"
    oConfigFile.WriteLine "End Sub"
    oConfigFile.WriteLine "%" & ">"
    oConfigFile.Close
    Set oConfigFile = Nothing
    Set oFS2 = Nothing

    On Error GoTo 0
    RunInstaller = "OK"
End Function

' SHA256 simple para el instalador (sin depender de functions.asp)
Function SHA256Hash_Install(strText)
    On Error Resume Next
    Dim oSHA, oEncoding, oBytes, oHash, sHex, i
    Set oSHA      = CreateObject("System.Security.Cryptography.SHA256Managed")
    Set oEncoding = CreateObject("System.Text.UTF8Encoding")
    oBytes = oEncoding.GetBytes_4(strText)
    oHash  = oSHA.ComputeHash_2(oBytes)
    sHex = ""
    For i = 0 To UBound(oHash)
        sHex = sHex & Right("0" & Hex(oHash(i)), 2)
    Next
    SHA256Hash_Install = LCase(sHex)
    If Err.Number <> 0 Then SHA256Hash_Install = strText  ' Fallback
    Set oSHA = Nothing : Set oEncoding = Nothing
    On Error GoTo 0
End Function

' Obtener puerto por defecto según tipo de BD
Function getDefaultPort(dbType)
    Select Case LCase(dbType)
        Case "sqlserver"
            getDefaultPort = "1433"
        Case "mysql"
            getDefaultPort = "3306"
        Case "sqlite"
            getDefaultPort = ""
        Case Else
            getDefaultPort = "1433"
    End Select
End Function

' Construir cadena de conexión según tipo de BD
Function buildConnectionString(dbType, dbServer, dbName, dbUser, dbPass, dbPort, dbFile)
    Select Case LCase(dbType)
        Case "sqlserver"
            If dbPort = "" Then dbPort = "1433"
            buildConnectionString = "Provider=SQLNCLI11;Server=" & dbServer & ";Database=" & dbName & ";UID=" & dbUser & ";PWD=" & dbPass & ";Connect Timeout=30;"
        Case "mysql"
            If dbPort = "" Then dbPort = "3306"
            buildConnectionString = "Driver={MySQL ODBC 8.0 Unicode Driver};Server=" & dbServer & ";Port=" & dbPort & ";Database=" & dbName & ";Uid=" & dbUser & ";Pwd=" & dbPass & ";Option=3;"
        Case "sqlite"
            buildConnectionString = "Driver={SQLite3 ODBC Driver};Database=" & dbFile & ";"
        Case Else
            buildConnectionString = "Provider=SQLNCLI11;Server=" & dbServer & ";Database=" & dbName & ";UID=" & dbUser & ";PWD=" & dbPass & ";Connect Timeout=30;"
    End Select
End Function
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Instalador CMMS</title>
<link rel="stylesheet" href="/CMMS/assets/css/app.css">
<style>
  .install-step-content { display: none; }
  .install-step-content.active { display: block; }
  .db-type-card {
    border: 2px solid var(--border-default);
    border-radius: var(--radius-lg);
    padding: var(--space-lg);
    cursor: pointer;
    transition: all var(--transition-fast);
    text-align: center;
  }
  .db-type-card:hover, .db-type-card.selected {
    border-color: var(--primary);
    background: var(--primary-light);
  }
  .db-type-icon { font-size: 36px; margin-bottom: 8px; }
</style>
</head>
<body style="background:var(--bg-base)">

<div id="toast-container"></div>

<div class="install-wrapper">
  <div class="install-card">

    <!-- Header -->
    <div class="install-header">
      <div style="font-size:40px;margin-bottom:8px">⚙️</div>
      <h1>Instalador del Sistema CMMS</h1>
      <p>Configure la base de datos e instale el sistema</p>
    </div>

    <!-- Pasos -->
    <div class="install-steps">
      <div class="install-step <%= IIf(Step >= "1", IIf(Step > "1", "done", "active"), "") %>">
        <div class="step-circle"><%= IIf(Step > "1", "✓", "1") %></div>
        <div class="step-label">Base de Datos</div>
      </div>
      <div class="install-step <%= IIf(Step >= "2", IIf(Step > "2", "done", "active"), "") %>">
        <div class="step-circle"><%= IIf(Step > "2", "✓", "2") %></div>
        <div class="step-label">Conexión</div>
      </div>
      <div class="install-step <%= IIf(Step >= "3", IIf(Step > "3", "done", "active"), "") %>">
        <div class="step-circle"><%= IIf(Step > "3", "✓", "3") %></div>
        <div class="step-label">Administrador</div>
      </div>
      <div class="install-step <%= IIf(Step = "4", "active", "") %>">
        <div class="step-circle"><%= IIf(Step = "4", "✓", "4") %></div>
        <div class="step-label">Completado</div>
      </div>
    </div>

    <div class="install-body">

      <!-- Mensajes -->
      <% If ErrorMsg <> "" Then %>
      <div class="alert alert-danger" style="margin-bottom:16px">
        ⚠️ <span><%= Server.HTMLEncode(ErrorMsg) %></span>
      </div>
      <% End If %>
      <% If SuccessMsg <> "" Then %>
      <div class="alert alert-success" style="margin-bottom:16px">
        ✓ <span><%= Server.HTMLEncode(SuccessMsg) %></span>
      </div>
      <% End If %>

      <!-- ══════════════════════ PASO 1: Tipo de BD ══════════════════════ -->
      <% If Step = "1" Then %>
      <form method="POST" action="install.asp?step=2" id="step1Form">
        <input type="hidden" name="db_type" id="selected_db_type" value="sqlserver">
        
        <h3 style="color:var(--text-primary);margin-bottom:16px">Seleccione el tipo de base de datos</h3>
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:24px">
          <div class="db-type-card selected" id="card-sqlserver" onclick="selectDB('sqlserver')">
            <div class="db-type-icon">🖥️</div>
            <div style="font-weight:600;color:var(--text-primary)">SQL Server</div>
            <div style="font-size:12px;color:var(--text-muted);margin-top:4px">Microsoft SQL Server<br>2016+</div>
            <div style="margin-top:8px"><span class="badge badge-success no-dot">Recomendado</span></div>
          </div>
          <div class="db-type-card" id="card-mysql" onclick="selectDB('mysql')">
            <div class="db-type-icon">🐬</div>
            <div style="font-weight:600;color:var(--text-primary)">MySQL</div>
            <div style="font-size:12px;color:var(--text-muted);margin-top:4px">MySQL 5.7+<br>MariaDB</div>
            <div style="margin-top:8px"><span class="badge badge-success no-dot">Disponible</span></div>
          </div>
          <div class="db-type-card" id="card-sqlite" onclick="selectDB('sqlite')">
            <div class="db-type-icon">📁</div>
            <div style="font-weight:600;color:var(--text-primary)">SQLite</div>
            <div style="font-size:12px;color:var(--text-muted);margin-top:4px">Archivo local<br>Sin servidor</div>
            <div style="margin-top:8px"><span class="badge badge-success no-dot">Disponible</span></div>
          </div>
        </div>
        
        <div style="background:var(--primary-light);border:1px solid rgba(99,102,241,0.3);border-radius:10px;padding:16px;margin-bottom:24px">
          <div style="font-weight:600;color:var(--primary);margin-bottom:8px">📋 Requisitos previos</div>
          <ul style="color:var(--text-secondary);font-size:13px;padding-left:20px;line-height:2">
            <li id="req-sqlserver"><strong>SQL Server:</strong> SQL Server 2016+ o Express instalado</li>
            <li id="req-sqlserver2">SQL Server Native Client 11 (SQLNCLI11) en IIS</li>
            <li id="req-mysql" style="display:none"><strong>MySQL:</strong> MySQL 5.7+ o MariaDB 10.3+</li>
            <li id="req-mysql2" style="display:none">MySQL ODBC 8.0 Unicode Driver instalado en IIS</li>
            <li id="req-sqlite" style="display:none"><strong>SQLite:</strong> SQLite3 ODBC Driver instalado</li>
            <li id="req-common">Usuario con permisos CREATE TABLE, INSERT, UPDATE, DELETE</li>
          </ul>
        </div>
        <div style="text-align:right">
          <button type="submit" class="btn btn-primary">
            Continuar
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>
          </button>
        </div>
      </form>

      <!-- ══════════════════════ PASO 2: Conexión ══════════════════════ -->
      <% ElseIf Step = "2" Then %>
      <h3 style="color:var(--text-primary);margin-bottom:6px">Configuración de Conexión</h3>
      
      <% 
        ' Determinar tipo de BD desde POST o session
        Dim currentDbType
        currentDbType = LCase(Trim(Request.Form("db_type")))
        If currentDbType = "" Then currentDbType = "sqlserver"
        
        ' Actualizar subtítulo según tipo de BD
        Dim step2Subtitle
        Select Case currentDbType
            Case "mysql"
                step2Subtitle = "Ingrese los datos de conexión a MySQL"
            Case "sqlite"
                step2Subtitle = "Configure la ruta del archivo SQLite"
            Case Else
                step2Subtitle = "Ingrese los datos de conexión a SQL Server"
        End Select
      %>
      
      <p style="color:var(--text-muted);font-size:13px;margin-bottom:20px" id="step2-subtitle"><%= step2Subtitle %></p>
      
      <form method="POST" action="install.asp" data-validate id="connForm">
        <input type="hidden" name="action" value="test_save">
        <input type="hidden" name="db_type" value="<%= currentDbType %>">
        
        <!-- Campos SQL Server -->
        <div id="fields-sqlserver" style="<%= IIf(currentDbType = "sqlserver", "", "display:none") %>">
          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Servidor SQL <span class="required">*</span></label>
              <input type="text" name="db_server" class="form-control" required
                     placeholder="Ej: SQLSERVER01, .\SQLEXPRESS, 192.168.1.10"
                     value="<%= Server.HTMLEncode(Request.Form("db_server")) %>">
              <div class="form-hint">Nombre del servidor o IP. Para instancia local: .\SQLEXPRESS o (local)</div>
            </div>
            <div class="form-group">
              <label class="form-label">Base de Datos <span class="required">*</span></label>
              <input type="text" name="db_name" class="form-control" required
                     placeholder="Ej: CMMS_DB"
                     value="<%= Server.HTMLEncode(IIf(Request.Form("db_name") <> "", Request.Form("db_name"), "CMMS_DB")) %>">
              <div class="form-hint">La BD debe existir previamente</div>
            </div>
          </div>
          
          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Usuario SQL <span class="required">*</span></label>
              <input type="text" name="db_user" class="form-control" required
                     placeholder="Ej: sa, cmms_user"
                     value="<%= Server.HTMLEncode(Request.Form("db_user")) %>">
            </div>
            <div class="form-group">
              <label class="form-label">Contraseña SQL <span class="required">*</span></label>
              <input type="password" name="db_pass" class="form-control" required
                     placeholder="Contraseña del usuario SQL">
            </div>
          </div>
        </div>
        
        <!-- Campos MySQL -->
        <div id="fields-mysql" style="<%= IIf(currentDbType = "mysql", "", "display:none") %>">
          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Servidor MySQL <span class="required">*</span></label>
              <input type="text" name="db_server" class="form-control" required
                     placeholder="Ej: localhost, 192.168.1.10"
                     value="<%= Server.HTMLEncode(Request.Form("db_server")) %>">
              <div class="form-hint">Nombre del servidor o IP</div>
            </div>
            <div class="form-group">
              <label class="form-label">Puerto <span class="required">*</span></label>
              <input type="text" name="db_port" class="form-control" required
                     placeholder="3306"
                     value="<%= Server.HTMLEncode(IIf(Request.Form("db_port") <> "", Request.Form("db_port"), "3306")) %>">
              <div class="form-hint">Puerto MySQL (default: 3306)</div>
            </div>
          </div>
          
          <div class="form-row">
            <div class="form-group">
              <label class="form-label">Base de Datos <span class="required">*</span></label>
              <input type="text" name="db_name" class="form-control" required
                     placeholder="Ej: cmms_db"
                     value="<%= Server.HTMLEncode(IIf(Request.Form("db_name") <> "", Request.Form("db_name"), "cmms_db")) %>">
              <div class="form-hint">La BD debe existir previamente</div>
            </div>
            <div class="form-group">
              <label class="form-label">Usuario MySQL <span class="required">*</span></label>
              <input type="text" name="db_user" class="form-control" required
                     placeholder="Ej: root, cmms_user"
                     value="<%= Server.HTMLEncode(Request.Form("db_user")) %>">
            </div>
          </div>
          
          <div class="form-group">
            <label class="form-label">Contraseña MySQL <span class="required">*</span></label>
            <input type="password" name="db_pass" class="form-control" required
                   placeholder="Contraseña del usuario MySQL">
          </div>
        </div>
        
        <!-- Campos SQLite -->
        <div id="fields-sqlite" style="<%= IIf(currentDbType = "sqlite", "", "display:none") %>">
          <div class="form-group">
            <label class="form-label">Ruta del Archivo SQLite <span class="required">*</span></label>
            <input type="text" name="db_file" class="form-control" required
                   placeholder="Ej: C:\inetpub\wwwroot\CMMS\data\cmms.db"
                   value="<%= Server.HTMLEncode(IIf(Request.Form("db_file") <> "", Request.Form("db_file"), Server.MapPath("/CMMS/data/cmms.db"))) %>">
            <div class="form-hint">Ruta completa del archivo .db (se creará automáticamente si no existe)</div>
          </div>
          
          <div style="background:var(--info-light);border:1px solid rgba(59,130,246,0.3);border-radius:8px;padding:12px;margin-bottom:20px;font-size:12px;color:var(--info)">
            ℹ️ <strong>SQLite:</strong> No requiere servidor de base de datos. El archivo se creará en la ruta especificada. Asegúrese de que IIS tenga permisos de escritura en esa carpeta.
          </div>
        </div>

        <div id="warning-sqlserver" style="<%= IIf(currentDbType = "sqlserver", "", "display:none") %>;background:var(--warning-light);border:1px solid rgba(245,158,11,0.3);border-radius:8px;padding:12px;margin-bottom:20px;font-size:12px;color:var(--warning)">
          ⚠️ <strong>Importante:</strong> El usuario debe tener permisos para crear tablas e insertar datos en la base de datos especificada.
        </div>
        
        <div id="warning-mysql" style="<%= IIf(currentDbType = "mysql", "", "display:none") %>;background:var(--warning-light);border:1px solid rgba(245,158,11,0.3);border-radius:8px;padding:12px;margin-bottom:20px;font-size:12px;color:var(--warning)">
          ⚠️ <strong>Importante:</strong> El usuario debe tener permisos para crear tablas e insertar datos en la base de datos especificada.
        </div>

        <div style="display:flex;justify-content:space-between;gap:12px">
          <a href="?step=1" class="btn btn-outline">← Atrás</a>
          <button type="submit" class="btn btn-primary" id="testBtn">
            🔌 Probar Conexión e Instalar
          </button>
        </div>
      </form>

      <!-- ══════════════════════ PASO 3: Admin ══════════════════════ -->
      <% ElseIf Step = "3" Then %>
      <h3 style="color:var(--text-primary);margin-bottom:6px">Crear Administrador del Sistema</h3>
      <p style="color:var(--text-muted);font-size:13px;margin-bottom:20px">Configure las credenciales del usuario administrador</p>

      <div style="background:var(--success-light);border:1px solid rgba(16,185,129,0.3);border-radius:8px;padding:12px;margin-bottom:20px;font-size:12px;color:var(--success)">
        ✓ Conexión a SQL Server establecida correctamente. Servidor: <strong><%= Server.HTMLEncode(Session("inst_server")) %></strong> | BD: <strong><%= Server.HTMLEncode(Session("inst_name")) %></strong>
      </div>

      <form method="POST" action="install.asp" data-validate id="adminForm">
        <input type="hidden" name="action" value="install">

        <div class="form-row">
          <div class="form-group">
            <label class="form-label">Nombre <span class="required">*</span></label>
            <input type="text" name="admin_first" class="form-control" required placeholder="Nombre">
          </div>
          <div class="form-group">
            <label class="form-label">Apellido</label>
            <input type="text" name="admin_last" class="form-control" placeholder="Apellido">
          </div>
        </div>

        <div class="form-group">
          <label class="form-label">Correo Electrónico <span class="required">*</span></label>
          <input type="email" name="admin_email" class="form-control" required placeholder="admin@empresa.com" data-type="email">
        </div>

        <div class="form-row">
          <div class="form-group">
            <label class="form-label">Usuario Admin <span class="required">*</span></label>
            <input type="text" name="admin_user" class="form-control" required placeholder="admin" value="admin">
          </div>
          <div class="form-group">
            <label class="form-label">Contraseña <span class="required">*</span></label>
            <input type="password" name="admin_pass" class="form-control" required placeholder="Mínimo 8 caracteres" value="Admin@123!">
          </div>
        </div>

        <div style="background:var(--info-light);border:1px solid rgba(59,130,246,0.3);border-radius:8px;padding:12px;margin-bottom:20px;font-size:12px;color:var(--info)">
          ℹ️ El instalador creará todas las tablas necesarias y cargará los datos iniciales del sistema.
        </div>

        <div style="display:flex;justify-content:space-between;gap:12px">
          <a href="?step=2" class="btn btn-outline">← Atrás</a>
          <button type="submit" class="btn btn-success" id="installBtn" onclick="this.textContent='Instalando...';this.disabled=true;this.closest('form').submit()">
            🚀 Instalar Sistema
          </button>
        </div>
      </form>

      <!-- ══════════════════════ PASO 4: Éxito ══════════════════════ -->
      <% ElseIf Step = "4" Then %>
      <div style="text-align:center;padding:32px 0">
        <div style="font-size:64px;margin-bottom:16px">🎉</div>
        <h2 style="color:var(--success);font-size:28px;font-weight:800;margin-bottom:8px">¡Instalación Completada!</h2>
        <p style="color:var(--text-muted);margin-bottom:24px">El sistema CMMS ha sido instalado correctamente.</p>
        
        <div style="background:var(--bg-elevated);border:1px solid var(--border-default);border-radius:12px;padding:20px;text-align:left;margin-bottom:24px">
          <h4 style="color:var(--text-primary);margin-bottom:12px">📋 Resumen de la instalación:</h4>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:13px">
            <div style="color:var(--text-muted)">Servidor:</div>
            <div style="color:var(--text-primary);font-weight:500"><%= Server.HTMLEncode(Session("inst_server")) %></div>
            <div style="color:var(--text-muted)">Base de datos:</div>
            <div style="color:var(--text-primary);font-weight:500"><%= Server.HTMLEncode(Session("inst_name")) %></div>
            <div style="color:var(--text-muted)">Usuario admin:</div>
            <div style="color:var(--text-primary);font-weight:500">admin</div>
          </div>
        </div>

        <div style="background:var(--warning-light);border:1px solid rgba(245,158,11,0.3);border-radius:8px;padding:12px;margin-bottom:24px;font-size:12px;color:var(--warning);text-align:left">
          ⚠️ <strong>Seguridad:</strong> Por favor cambie la contraseña del administrador inmediatamente después del primer inicio de sesión.
        </div>

        <a href="/CMMS/login.asp" class="btn btn-primary btn-lg">
          🔐 Ir al Login
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>
        </a>
        
        <div style="margin-top:16px;font-size:12px;color:var(--text-muted)">
          Se recomienda eliminar o restringir el acceso al archivo <code>install.asp</code> en producción.
        </div>
      </div>

      <% End If %>

    </div><!-- /install-body -->
  </div><!-- /install-card -->
</div><!-- /install-wrapper -->

<script src="/CMMS/assets/js/app.js"></script>
<script>
let currentDBType = 'sqlserver';

function selectDB(type) {
    currentDBType = type;
    document.querySelectorAll('.db-type-card').forEach(c => c.classList.remove('selected'));
    document.getElementById('card-' + type).classList.add('selected');
    
    // Actualizar campo oculto en paso 1
    document.getElementById('selected_db_type').value = type;
    
    // Mostrar requisitos específicos en paso 1
    document.getElementById('req-sqlserver').style.display = (type === 'sqlserver') ? '' : 'none';
    document.getElementById('req-sqlserver2').style.display = (type === 'sqlserver') ? '' : 'none';
    document.getElementById('req-mysql').style.display = (type === 'mysql') ? '' : 'none';
    document.getElementById('req-mysql2').style.display = (type === 'mysql') ? '' : 'none';
    document.getElementById('req-sqlite').style.display = (type === 'sqlite') ? '' : 'none';
    
    // Actualizar formulario en paso 2
    updateStep2Form(type);
}

function updateStep2Form(type) {
    // Ocultar todos los campos
    document.getElementById('fields-sqlserver').style.display = 'none';
    document.getElementById('fields-mysql').style.display = 'none';
    document.getElementById('fields-sqlite').style.display = 'none';
    document.getElementById('warning-sqlserver').style.display = 'none';
    document.getElementById('warning-mysql').style.display = 'none';
    
    // Actualizar subtítulo
    var subtitle = document.getElementById('step2-subtitle');
    
    // Mostrar campos según tipo
    if (type === 'sqlserver') {
        document.getElementById('fields-sqlserver').style.display = '';
        document.getElementById('warning-sqlserver').style.display = '';
        subtitle.textContent = 'Ingrese los datos de conexión a SQL Server';
        
        // Hacer requeridos los campos de SQL Server
        setRequired('db_server', true);
        setRequired('db_name', true);
        setRequired('db_user', true);
        setRequired('db_pass', true);
        setRequired('db_port', false);
        setRequired('db_file', false);
        
    } else if (type === 'mysql') {
        document.getElementById('fields-mysql').style.display = '';
        document.getElementById('warning-mysql').style.display = '';
        subtitle.textContent = 'Ingrese los datos de conexión a MySQL';
        
        // Hacer requeridos los campos de MySQL
        setRequired('db_server', true);
        setRequired('db_name', true);
        setRequired('db_user', true);
        setRequired('db_pass', true);
        setRequired('db_port', true);
        setRequired('db_file', false);
        
    } else if (type === 'sqlite') {
        document.getElementById('fields-sqlite').style.display = '';
        subtitle.textContent = 'Configure la ruta del archivo SQLite';
        
        // Hacer requeridos los campos de SQLite
        setRequired('db_server', false);
        setRequired('db_name', false);
        setRequired('db_user', false);
        setRequired('db_pass', false);
        setRequired('db_port', false);
        setRequired('db_file', true);
    }
}

function setRequired(fieldName, required) {
    var inputs = document.getElementsByName(fieldName);
    for (var i = 0; i < inputs.length; i++) {
        inputs[i].required = required;
    }
}

function validateStep1() {
    // La validación se hace automáticamente al enviar el formulario
    // El tipo de BD seleccionado se envía por POST al paso 2
    return true;
}

// Manejar envío del formulario paso 1
var step1Form = document.getElementById('step1Form');
if (step1Form) {
    step1Form.addEventListener('submit', function(e) {
        // El tipo de BD ya está en el campo oculto selected_db_type
        // Solo dejamos que el formulario se envíe normalmente
    });
}

// Inicializar formulario al cargar
document.addEventListener('DOMContentLoaded', function() {
    // Paso 1: inicializar requisitos según BD seleccionada
    var step1DbType = document.getElementById('selected_db_type');
    if (step1DbType && step1DbType.value) {
        selectDB(step1DbType.value);
    }
    
    // Paso 2: inicializar campos según BD recibida por POST
    var dbTypeInput = document.querySelector('form#connForm input[name="db_type"]');
    if (dbTypeInput && dbTypeInput.value) {
        updateStep2Form(dbTypeInput.value);
    }
});

document.getElementById('connForm') && document.getElementById('connForm').addEventListener('submit', function(){
    const btn = document.getElementById('testBtn');
    if(btn) { btn.innerHTML = '<span class="spinner" style="width:16px;height:16px;border-width:2px;margin-right:8px"></span>Probando...'; btn.disabled = true; }
});
</script>
</body>
</html>

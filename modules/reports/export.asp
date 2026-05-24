<!--#include virtual="/CMMS/config/database.asp"-->
<!--#include virtual="/CMMS/core/functions.asp"-->
<!--#include virtual="/CMMS/core/auth.asp"-->
<%
CheckAuth()

Dim exportType : exportType = Request.QueryString("type")
Dim oConn : Set oConn = GetConnection()
Dim rs, sql

Response.ContentType = "text/csv"
Response.AddHeader "Content-Disposition", "attachment; filename=export_" & exportType & "_" & Year(Now)&Month(Now)&Day(Now) & ".csv"

Select Case exportType
    Case "wo"
        sql = "SELECT wo.code, wo.title, wo.type, wo.priority, wo.status, p.name AS plant_name, a.code AS asset_code, wo.scheduled_start, wo.created_at " & _
              "FROM cmms_work_orders wo " & _
              "LEFT JOIN cmms_plants p ON p.id = wo.plant_id " & _
              "LEFT JOIN cmms_assets a ON a.id = wo.asset_id " & _
              "ORDER BY wo.created_at DESC"
        
        Set rs = oConn.Execute(sql)
        Response.Write "Codigo,Titulo,Tipo,Prioridad,Estado,Planta,Equipo,Programado,Fecha Creacion" & vbCrLf
        
        Do While Not rs.EOF
            Response.Write """" & rs("code") & """,""" & Replace(rs("title"), """", """""") & """,""" & rs("type") & """,""" & rs("priority") & """,""" & rs("status") & """,""" & rs("plant_name") & """,""" & rs("asset_code") & """,""" & rs("scheduled_start") & """,""" & rs("created_at") & """" & vbCrLf
            rs.MoveNext
        Loop
        rs.Close : Set rs = Nothing

    Case "inventory"
        sql = "SELECT i.code, i.name, i.category, p.name AS plant_name, i.unit_cost, i.reorder_point, ISNULL((SELECT SUM(quantity) FROM cmms_inventory_stock WHERE inventory_id = i.id), 0) AS stock " & _
              "FROM cmms_inventory i " & _
              "LEFT JOIN cmms_plants p ON p.id = i.plant_id " & _
              "ORDER BY i.name"
              
        Set rs = oConn.Execute(sql)
        Response.Write "Codigo,Articulo,Categoria,Planta,CostoUnitario,PuntoReorden,StockActual" & vbCrLf
        
        Do While Not rs.EOF
            Response.Write """" & rs("code") & """,""" & Replace(rs("name"), """", """""") & """,""" & rs("category") & """,""" & rs("plant_name") & """,""" & rs("unit_cost") & """,""" & rs("reorder_point") & """,""" & rs("stock") & """" & vbCrLf
            rs.MoveNext
        Loop
        rs.Close : Set rs = Nothing

    Case "assets"
        sql = "SELECT a.code, a.name, a.category, p.name AS plant_name, a.criticality, a.status " & _
              "FROM cmms_assets a " & _
              "LEFT JOIN cmms_plants p ON p.id = a.plant_id " & _
              "ORDER BY a.name"
              
        Set rs = oConn.Execute(sql)
        Response.Write "Codigo,Equipo,Categoria,Planta,Criticidad,Estado" & vbCrLf
        
        Do While Not rs.EOF
            Response.Write """" & rs("code") & """,""" & Replace(rs("name"), """", """""") & """,""" & rs("category") & """,""" & rs("plant_name") & """,""" & rs("criticality") & """,""" & rs("status") & """" & vbCrLf
            rs.MoveNext
        Loop
        rs.Close : Set rs = Nothing

    Case Else
        Response.Write "Tipo de exportación no válido."
End Select

CloseConnection(oConn)
Response.End
%>

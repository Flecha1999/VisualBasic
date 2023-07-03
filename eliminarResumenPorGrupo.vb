Sub Main
	Stop
	Set xview = NewCompoundView( self, "GRUPOEMPLEADOS", self.Workspace, nil, true )
	xview.addbocol("Nombre")
	xview.ADDORDERCOLUMN(NewOrderSpec( NewColumnSpec( "GRUPOEMPLEADOS", "Nombre", "GRUPOEMPLEADOS" ), FALSE ))
	
	Set xVisualVar = VisualVarEditor( "Seleccionar Grupo" )	
	Call AddVarView   ( xVisualVar, "Ingrese Grupo", "Grupo","Grupo" , xview, "Nombre")
	xAcept = ShowVisualVar( xVisualVar )
	If xacept then
	   xGrupo = GetValueVisualVar( xVisualVar, "Ingrese Grupo", "Grupo" )
	   Set oGrupo = instanciarbo(xGrupo, "GRUPOEMPLEADOS", self.workspace)
	   For Each oEmpleado in oGrupo.empleados
	   	   Set xview = NewCompoundView( self, "RESUMENLIQUIDACION", self.Workspace, nil, true )
		   xView.addfilter( NewFilterSpec(NewColumnSpec( "RESUMENLIQUIDACION", "estado", "RESUMENLIQUIDACION" ),"=",  "0" ))
		   xView.addfilter( NewFilterSpec(NewColumnSpec( "RESUMENLIQUIDACION", "EMPLEADO", "RESUMENLIQUIDACION" ),"=",  oEmpleado.id ))
		   xView.addfilter( NewFilterSpec(NewColumnSpec( "RESUMENLIQUIDACION", "LIQUIDACION", "RESUMENLIQUIDACION" ),"=",  self.id ))		   
		   If Not xView.ViewItems.IsEmpty then
		   	  Set oResumen  = xView.ViewItems.first.current.bo
			  SendDebug "Eliminando Resumen ( " & oResumen.Liquidacion.TipoLiquidacion.Codigo & " ) del empleado: " & oResumen.Legajo & " - " & oResumen.NombreEmpleado
			  oResumen.delete
			  Call workspacecheck(self.workspace)
			  SendDebug "Resumen Eliminado"
	       End If
	   Next
	   MsgBox "Proceso Fiinalizado"
    Else
	   Call msgbox("Proceso Cancelado por el Usuario.")
	End IF

End Sub

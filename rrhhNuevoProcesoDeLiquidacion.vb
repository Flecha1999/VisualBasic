' IC - Fx Procesar Liquidacion ( Nuevo )
Sub Main
	Stop
	If UCASE(nombreusuario) = "DESARROLLO2" OR PerteneceAGrupo( "CALIPSO RH LIQUIDACIONES" ) Then
		Path = ""
		nCentroCostos = ""
		Set xEsquema  = Self.Place.Owner.UnidadOperativa.EsquemaOperativo
		
		Set xViewCC  = NewCompoundView( Self, "CENTROCOSTOS", Self.Workspace, Nil, True )
		xViewCC.addjoin(NewJoinSpec( NewColumnSpec( "CENTROCOSTOS", "UNIDADOPERATIVA", "" ), NewColumnSpec( "UOSOLICITUD", "ID", "" ), False ))
 		xViewCC.addFilter(NewFilterSpec(NewColumnSpec("CENTROCOSTOS", "ACTIVESTATUS", ""), "=", 0))
		xViewCC.addOrderColumn(NewOrderSpec(NewColumnSpec("CENTROCOSTOS", "NOMBRE", ""), True ))
		xViewCC.ColumnFromPath( "NOMBRE" )
		
		Set xViewCN  = NewCompoundView( Self, "PERFIL", Self.Workspace, Nil, True )
		xViewCN.addOrderColumn(NewOrderSpec(NewColumnSpec("PERFIL", "DESCRIPCION", ""), True ))
		xViewCN.ColumnFromPath( "DESCRIPCION" )
		
		Set xViewCA  = NewCompoundView( Self, "CATEGORIA", Self.Workspace, Nil, True )
		xViewCA.addOrderColumn(NewOrderSpec(NewColumnSpec("CATEGORIA", "NOMBRE", ""), True ))
		xViewCA.ColumnFromPath( "NOMBRE" )
		
		Set xViewGE  = NewCompoundView( Self, "GRUPOEMPLEADOS", Self.Workspace, Nil, True )
		xViewGE.addOrderColumn(NewOrderSpec(NewColumnSpec("GRUPOEMPLEADOS", "NOMBRE", ""), True ))
		xViewGE.ColumnFromPath( "NOMBRE" )
		
		Set xViewEM = NewCompoundView( Self, "EMPLEADO", Self.Workspace, Nil, True )
		xViewEM.addjoin(NewJoinSpec( NewColumnSpec( "EMPLEADO", "UNIDADOPERATIVA", "" ), NewColumnSpec( "UORECURSOSHUMANOS", "ID", "" ), False ))
		xViewEM.addFilter(NewFilterSpec(NewColumnSpec("EMPLEADO", "ACTIVESTATUS", ""), "=", 0))
		xViewEM.addOrderColumn(NewOrderSpec(NewColumnSpec("EMPLEADO", "DESCRIPCION", ""), True ))
		xViewEM.ColumnFromPath( "DESCRIPCION" )
				
		Set xvisualvar = VisualVarEditor( "Proceso de Liquidacion" )
		Call AddVarView( xvisualvar, "01_dEmpleado", "Desde Empleado"  , "1 - Rango Empleados" , xViewEM, "DESCRIPCION" )
		Call AddVarView( xvisualvar, "02_hEmpleado", "Hasta Empleado"  , "1 - Rango Empleados" , xViewEM, "DESCRIPCION" )
		Call AddVarView( xvisualvar, "04_CN"       , "Convenio"        , "2 - Más Filtros"     , xViewCN, "DESCRIPCION" )
		Call AddVarView( xvisualvar, "03_cc"       , "Centro de Costos", "2 - Más Filtros"	   , xViewCC, "NOMBRE" )
		Call AddVarView( xVisualVar, "05_categoria", "Categoria"       , "2 - Más Filtros"     , xViewCA, "NOMBRE")
		Call AddVarView( xVisualVar, "06_Grupo"    , "Grupo"           , "2 - Más Filtros"     , xViewGE, "NOMBRE")
		Do
			OK = True 
			MENSAJE_ERROR = ""
			Set oCentroCostos = Nothing
			Set oTipoLiquidacion = Nothing
			xacept = ShowVisualVar( xvisualvar )
			If xacept Then
				xEmpleadoDesde	= GetValueVisualVar( xvisualvar, "01_dEmpleado", "1 - Rango Empleados")
				xEmpleadoHasta	= GetValueVisualVar( xvisualvar, "02_hEmpleado", "1 - Rango Empleados")
				xConvenio		= GetValueVisualVar( xVisualVar, "04_CN"       , "2 - Más Filtros")
				xCentroCostos	= GetValueVisualVar( xvisualvar, "03_cc"       , "2 - Más Filtros")
				xCategoria		= GetValueVisualVar( xVisualVar, "05_categoria", "2 - Más Filtros")
				xGrupoEmpleados	= GetValueVisualVar( xVisualVar, "06_Grupo"    , "2 - Más Filtros")
				nCentroCostos	= ""
				nConvenio		= ""
				
				If IsEmpty(xCentroCostos) And IsEmpty(xConvenio) And IsEmpty(xCategoria) And IsEmpty(xGrupoEmpleados) And (IsEmpty(xEmpleadoDesde) Or IsEmpty(xEmpleadoHasta)) Then
					OK = FALSE : MENSAJE_ERROR = "No se definieron correctamente los parámetros." & chr(13)
				Else
				    
				End If
				If Not OK Then
					MsgBox MENSAJE_ERROR, 16, "Error" : MENSAJE_ERROR = ""
				End If
			Else
'			   MsgBox "Proceso cancelado por el usuario", 64, "Proceso de Liquidacion"
			   Exit Sub
			End If
		Loop While Not OK	

		Set xLiquidacion   = Self
		Set xViewEM = NewCompoundView( Self, "EMPLEADO", Self.Workspace, Nil, True )
		xViewEM.addjoin(NewJoinSpec( NewColumnSpec( "EMPLEADO", "UNIDADOPERATIVA", "" ), NewColumnSpec( "UORECURSOSHUMANOS", "ID", "" ), False ))
        xViewEM.addjoin(NewJoinSpec( NewColumnSpec( "EMPLEADO", "BOEXTENSION", "" ), NewColumnSpec( "UD_EMPLEADO", "ID", "" ), False ))
		
		If Not IsEmpty(xEmpleadoDesde) And Not IsEmpty(xEmpleadoHasta) Then
			If Not IsEmpty(xEmpleadoDesde) Then
				Set oEmpleadoDesde = ExisteBo(Self, "EMPLEADO", "id", xEmpleadoDesde, nil, True, false, "=")
				xViewEM.addFilter(NewFilterSpec(NewColumnSpec("EMPLEADO", "CODIGO", ""), ">=", oEmpleadoDesde.Codigo))
			End If
			If Not IsEmpty(xEmpleadoHasta) Then
				Set oEmpleadoHasta = ExisteBo(Self, "EMPLEADO", "id", xEmpleadoHasta, nil, True, false, "=")
				xViewEM.addFilter(NewFilterSpec(NewColumnSpec("EMPLEADO", "CODIGO", ""), "<=", oEmpleadoHasta.Codigo))
			End If
		End If
		
		If Not IsEmpty(xCentroCostos) Then
			Set oCentroCostos = ExisteBo(Self, "CENTROCOSTOS", "id", xCentroCostos, nil, True, false, "=")
			xViewEM.addFilter(NewFilterSpec(NewColumnSpec("EMPLEADO", "CENTROCOSTOS", ""), "=", xCentroCostos))
		End If
		
		If Not IsEmpty(xConvenio) Then
			Set oConvenio = ExisteBo(Self, "PERFIL", "id", xConvenio, nil, True, false, "=")
			xViewEM.addFilter(NewFilterSpec(NewColumnSpec("EMPLEADO", "PERFIL", ""), "=", xConvenio))
		End If
		
		If Not IsEmpty(xCategoria) Then
			Set oCategoria = ExisteBo(Self, "CATEGORIA", "id", xCategoria, nil, True, false, "=")
			xViewEM.addFilter(NewFilterSpec(NewColumnSpec("EMPLEADO", "CATEGORIA", ""), "=", xCategoria))
		End If

		If Not IsEmpty(xGrupoEmpleados) Then
			Set oGrupoEmpleados = ExisteBo(Self, "GRUPOEMPLEADOS", "id", xGrupoEmpleados, nil, True, false, "=")
			xViewEM.addjoin(NewJoinSpec( NewColumnSpec( "EMPLEADO", "ID", "" ), NewColumnSpec( "PERSLIST", "ITEM", "" ), False ))
			xViewEM.addjoin(NewJoinSpec( NewColumnSpec( "PERSLIST", "ID", "" ), NewColumnSpec( "BOLIST", "BO_ITEMS", "" ), False ))
			xViewEM.addjoin(NewJoinSpec( NewColumnSpec( "BOLIST", "ID", "" ), NewColumnSpec( "GRUPOEMPLEADOS", "EMPLEADOS", "" ), False ))
			xViewEM.addFilter(NewFilterSpec(NewColumnSpec("GRUPOEMPLEADOS", "ID", ""), "=", oGrupoEmpleados.Id))
		End If
		' El sistema NO liquida personas que posean fecha de alta posterior a la fecha de cierre de la liquidación
		xViewEM.addFilter(NewFilterSpec(NewColumnSpec("EMPLEADO", "FECHAINGRESO", ""), "<=", xLiquidacion.FinPeriodo))
		
		' ----- Validaciones - Empleados que no se Liquidaran...  ----- '
        xCantEmpInhabilitados = 0 : xEmpInhabilitados = ""
		For Each oRec In xViewEM.ViewItems
			Set oEmp = oRec.Bo
			If Not oEmp.BoExtension.HabilitadoLiquidacion Then
                xCantEmpInhabilitados = xCantEmpInhabilitados + 1
                If xCantEmpInhabilitados = 1 Then
                    xEmpInhabilitados = oEmp.Descripcion
                Else
                    xEmpInhabilitados = xEmpInhabilitados & Chr(13) & oEmp.Descripcion
                End If
            End If
		Next
        xViewEM.addFilter(NewFilterSpec(NewColumnSpec("UD_EMPLEADO", "HABILITADOLIQUIDACION", ""), " = ", True))

		'Call EscribirTXT(path,xString,1)		
		' ---------------------------------------------------------------- '
		
		' El sistema NO liquida personas que posean motivo de baja, por mas que esten activos
'		xViewEM.addFilter(NewFilterSpec(NewColumnSpec("EMPLEADO", "MOTIVOBAJA", ""), " IS ", NULL))
		
		Liquidar    = False
		FechaInicio = CDate(xLiquidacion.InicioPeriodo)
		FechaHasta  = CDate(xLiquidacion.FinPeriodo)
		xMensaje    = ""
		Select Case xLiquidacion.TipoLiquidacion.Codigo
			Case "07-FIN"	'	Final
				For Each oItem In xViewEM.ViewItems
					Select Case oItem.Bo.PERFIL.ID
						Case "{D86EE5BA-6F25-4525-9803-5F4FEAA852C9}" : xConcepto = "225"    ' Comercio
						Case "{A9714C75-98F1-40EE-B9C4-075A5916587E}" : xConcepto = "1821"   ' Pasantes
						Case "{8A8D8A26-03BF-422D-86E4-69F6E0E0DF62}" : xConcepto = "1380"   ' Sorb
						Case "{75D7224D-06AF-423F-802C-76B86430179F}" : xConcepto = "860"    ' Smata
						Case "{495CED3E-3998-44B1-88B3-1785BBDC93BC}" : xConcepto = "542"    ' Soelsac
						Case "{8E6C3E23-2EA1-47C9-808D-14DD8E591718}" : xConcepto = "1590"   ' Fuera de convenio
						Case "{68344A7D-4055-4A30-8881-71ADB7EC707F}" : xConcepto = "1731"   ' Parque y jardines
						Case "{BDD8FB00-8B97-4C7E-A45A-C637543598B3}" : xConcepto = "2124"   ' U.O.M.
						Case "{CF65CADD-83AA-43BF-AED4-CF7071A2A061}" : xConcepto = "1099"   ' S.O.M
						Case "{6DC95932-42F4-40A1-ABE3-69A680AA7A5A}" : xConcepto = "1065"   'SAC SOM
					End Select
					xTotalHaberM = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "01-MES", "A", FechaInicio, FechaHasta, "SUM")
					If xTotalHaberM > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un recibo mensual liquidado." 
					End If
					xTotalHaberV = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "05-VAC", "A", FechaInicio, FechaHasta, "SUM")
					If xTotalHaberV > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un LIQ. VACACIONES liquidado."
					 End If
				Next
			Case "02-1aQ"	'	Primera quincena
				For Each oItem In xViewEM.ViewItems
					Select Case oItem.Bo.PERFIL.ID
						Case "{D86EE5BA-6F25-4525-9803-5F4FEAA852C9}" : xConcepto = "225"    ' Comercio
						Case "{A9714C75-98F1-40EE-B9C4-075A5916587E}" : xConcepto = "1821"   ' Pasantes
						Case "{8A8D8A26-03BF-422D-86E4-69F6E0E0DF62}" : xConcepto = "1380"   ' Sorb
						Case "{75D7224D-06AF-423F-802C-76B86430179F}" : xConcepto = "860"    ' Smata
						Case "{495CED3E-3998-44B1-88B3-1785BBDC93BC}" : xConcepto = "542"    ' Soelsac
						Case "{8E6C3E23-2EA1-47C9-808D-14DD8E591718}" : xConcepto = "1590"   ' Fuera de convenio
						Case "{68344A7D-4055-4A30-8881-71ADB7EC707F}" : xConcepto = "1731"   ' Parque y jardines
						Case "{BDD8FB00-8B97-4C7E-A45A-C637543598B3}" : xConcepto = "2124"   ' U.O.M.
						Case "{CF65CADD-83AA-43BF-AED4-CF7071A2A061}" : xConcepto = "1099"   ' S.O.M
						Case "{6DC95932-42F4-40A1-ABE3-69A680AA7A5A}" : xConcepto = "1065"   'SAC SOM
					End Select
					xTotalHaberM = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "01-MES", "A", FechaInicio, FechaHasta, "SUM")
					If xTotalHaberM > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un recibo mensual liquidado." 
					End If
					xTotalHaberV = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "05-VAC", "A", FechaInicio, FechaHasta, "SUM")
					If xTotalHaberV > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un LIQ. VACACIONES liquidado."
					 End If
				Next
			Case "03-2aQ"	'	Segunda quincena
				For Each oItem In xViewEM.ViewItems
					Select Case oItem.Bo.PERFIL.ID
						Case "{D86EE5BA-6F25-4525-9803-5F4FEAA852C9}" : xConcepto = "225"    ' Comercio
						Case "{A9714C75-98F1-40EE-B9C4-075A5916587E}" : xConcepto = "1821"   ' Pasantes
						Case "{8A8D8A26-03BF-422D-86E4-69F6E0E0DF62}" : xConcepto = "1380"   ' Sorb
						Case "{75D7224D-06AF-423F-802C-76B86430179F}" : xConcepto = "860"    ' Smata
						Case "{495CED3E-3998-44B1-88B3-1785BBDC93BC}" : xConcepto = "542"    ' Soelsac
						Case "{8E6C3E23-2EA1-47C9-808D-14DD8E591718}" : xConcepto = "1590"   ' Fuera de convenio
						Case "{68344A7D-4055-4A30-8881-71ADB7EC707F}" : xConcepto = "1731"   ' Parque y jardines
						Case "{BDD8FB00-8B97-4C7E-A45A-C637543598B3}" : xConcepto = "2124"   ' U.O.M.
						Case "{CF65CADD-83AA-43BF-AED4-CF7071A2A061}" : xConcepto = "1099"   ' S.O.M
						Case "{6DC95932-42F4-40A1-ABE3-69A680AA7A5A}" : xConcepto = "1065"   'SAC SOM
					End Select
					xTotalHaberM = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "01-MES", "A", FechaInicio, FechaHasta, "SUM")
					If xTotalHaberM > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un recibo mensual liquidado." 
					End If
					xTotalHaberV = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "05-VAC", "A", FechaInicio, FechaHasta, "SUM")
					If xTotalHaberV > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un LIQ. VACACIONES liquidado."
					 End If
				Next

			Case "09-ACU"	'	Liquidacion Acuerdos
				For Each oItem In xViewEM.ViewItems
					xTotalHaberM = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "01-MES", "A", FechaInicio, FechaHasta, "SUM")
					If xTotalHaberM > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un LIQ. MENSUAL cargada." 
					End If
					xTotalHaberV = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "05-VAC", "A", FechaInicio, FechaHasta, "SUM")
					If xTotalHaberV > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un LIQ. VACACIONES cargada." 
					End If
				Next
			Case "01-MES"	'	Mensual
				For Each oItem In xViewEM.ViewItems
					Select Case oItem.Bo.PERFIL.ID
						Case "{D86EE5BA-6F25-4525-9803-5F4FEAA852C9}" : xConcepto = "225"    ' Comercio
						Case "{A9714C75-98F1-40EE-B9C4-075A5916587E}" : xConcepto = "1821"   ' Pasantes
						Case "{8A8D8A26-03BF-422D-86E4-69F6E0E0DF62}" : xConcepto = "1380"   ' Sorb
						Case "{75D7224D-06AF-423F-802C-76B86430179F}" : xConcepto = "860"    ' Smata
						Case "{495CED3E-3998-44B1-88B3-1785BBDC93BC}" : xConcepto = "542"    ' Soelsac
						Case "{8E6C3E23-2EA1-47C9-808D-14DD8E591718}" : xConcepto = "1590"   ' Fuera de convenio
						Case "{68344A7D-4055-4A30-8881-71ADB7EC707F}" : xConcepto = "1731"   ' Parque y jardines
						Case "{BDD8FB00-8B97-4C7E-A45A-C637543598B3}" : xConcepto = "2124"   ' U.O.M.
						Case "{CF65CADD-83AA-43BF-AED4-CF7071A2A061}" : xConcepto = "1099"   ' S.O.M
						Case "{6DC95932-42F4-40A1-ABE3-69A680AA7A5A}" : xConcepto = "1065"   'SAC SOM
					End Select
					xTotalHaberF = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "07-FIN", "A", FechaInicio, FechaHasta, "SUM")
					If xTotalHaberF > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un LIQ. FINAL cargada." 
					End If
					xTotalHaberA = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "09-ACU", "A", FechaInicio, FechaHasta, "SUM")
					If xTotalHaberA > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un LIQ. ACUERDO cargada." 
					End If
				Next
			Case "05-VAC"	'	Vacaciones
				For Each oItem In xViewEM.ViewItems
					Select Case oItem.Bo.PERFIL.ID
						Case "{D86EE5BA-6F25-4525-9803-5F4FEAA852C9}" : xConcepto = "225"    ' Comercio
						Case "{A9714C75-98F1-40EE-B9C4-075A5916587E}" : xConcepto = "1821"   ' Pasantes
						Case "{8A8D8A26-03BF-422D-86E4-69F6E0E0DF62}" : xConcepto = "1380"   ' Sorb
						Case "{75D7224D-06AF-423F-802C-76B86430179F}" : xConcepto = "860"    ' Smata
						Case "{495CED3E-3998-44B1-88B3-1785BBDC93BC}" : xConcepto = "542"    ' Soelsac
						Case "{8E6C3E23-2EA1-47C9-808D-14DD8E591718}" : xConcepto = "1590"   ' Fuera de convenio
						Case "{68344A7D-4055-4A30-8881-71ADB7EC707F}" : xConcepto = "1731"   ' Parque y jardines
						Case "{BDD8FB00-8B97-4C7E-A45A-C637543598B3}" : xConcepto = "2124"   ' U.O.M.
						Case "{CF65CADD-83AA-43BF-AED4-CF7071A2A061}" : xConcepto = "1099"   ' S.O.M
						Case "{6DC95932-42F4-40A1-ABE3-69A680AA7A5A}" : xConcepto = "1065"   'SAC SOM
					End Select
					xTotalHaberF = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "07-FIN", "A", FechaInicio, FechaHasta, "SUM")'corregior 1099
					If xTotalHaberF > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un LIQ. FINAL cargada."
					End If
					xTotalHaberA = BuscarConceptoHistorico3(xConcepto, "I", oItem.Bo, "09-ACU", "A", FechaInicio, FechaHasta, "SUM")
					If xTotalHaberA > 0 Then 
						xMensaje = xMensaje & Chr(13) & "El empleado: " & oItem.Bo.Descripcion & ", tiene un LIQ. ACUERDO cargada." 
					End If
				Next
		End Select
'		If xMensaje <> "" Then
'			Call MsgBox ("Observaciones:" & Chr(13) & xMensaje & Chr(13) & Chr(13) & "Proceso Cancelado.", 64, "Información")
'			Exit Sub
'		End If
        'agg yo. probar
		'If xLiquidacion.Estado = 0 And month(Date) = xLiquidacion.MES And xLiquidacion.ANIO = Year(Date) Then ' 0 - Liquidacion abierta de este mes, de este año
		'	MsgBox xLiquidacion.Name + " no se puede liquidar porque esta abierta y no eliminada"
		'End If
		' Se agregan todos los filtros necesarios a la vista de empleados
		If xLiquidacion.Estado = 0 Then ' 0 - Liquidacion Abierta
			Call ProcesarLiquidacion(xLiquidacion, xViewEM)
            If xCantEmpInhabilitados > 0 Then
'               Call MsgBox ("Empleados Inhabilitados:  " & xCantEmpInhabilitados & Chr(13) & Chr(13) & xEmpInhabilitados, 64, "Información")
				xMsg = "Empleados Inhabilitados:  " & xCantEmpInhabilitados & Chr(13) & Chr(13) & xEmpInhabilitados
            End If
			If xMensaje <> "" Then
			    xMsg = xMsg & Chr(13) & Chr(13) & "Observaciones:" & Chr(13) & xMensaje & Chr(13) & Chr(13)
			End If
			If xMsg <> "" Then
			    Call MsgBox (xMsg, 64, "Información")
			End If
		Else
			MsgBox xLiquidacion.Name + " no se puede liquidar porque esta cerrada"
		End If
	End If
End Sub

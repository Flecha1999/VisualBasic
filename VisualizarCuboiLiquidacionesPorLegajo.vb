' Cubo Exportar a Calipso Analyzer - Analyzer Cubo de Liquidaciones x Legajo
'Sub main()
'    Set oCuboDef = ExisteBO(Self, "BOSQLCUBELAYOUTDEF", "LAYOUTNAME", "Analyzer Cubo de Liquidaciones x Legajo", nil, True, False, "=")
'    Set xDic = NewDic
'	' -------------------------------------------------------------------- '
'    xFechaDesde = dateserial(year(now()), month(now()), 1 )		  ' dateserial(year(now()), month(now)-1, 1 ))
'    xFechaHasta = dateserial(year(now()), month(now())+1, 1 )-1	  ' dateserial(year(now()), month(now), 1 )-1)
'    ' -------------------------------------------------------------------- '
'	Set xVisualvar = VisualVarEditor( "Detalle de Liquidaciones" )
'	Call AddVarDate( xvisualvar, "01_dfecha", "Desde Fecha", "Periodo de Tiempo", xFechaDesde)  ' dateserial(year(now()), month(now), 1 ))		  ' dateserial(year(now()), month(now)-1, 1 ))
'	Call AddVarDate( xvisualvar, "02_hfecha", "Hasta Fecha", "Periodo de Tiempo", xFechaHasta)  ' dateserial(year(now()), month(now)+1, 1 )-1)	  ' dateserial(year(now()), month(now), 1 )-1)
'	xAcept = ShowVisualVar( xvisualvar )
'	If xAcept Then
'        xdesdefecha   = cdate(GetValueVisualVar( xvisualvar, "01_dfecha", "Periodo de Tiempo"))
'        xhastafecha   = cdate(GetValueVisualVar( xvisualvar, "02_hfecha", "Periodo de Tiempo"))
'        strdesdefecha = year(xdesdefecha) & right("00" & month(xdesdefecha),2) & right("00" & day(xdesdefecha),2) 
'        strhastafecha = year(xhastafecha) & right("00" & month(xhastafecha),2) & right("00" & day(xhastafecha),2)
'        RegistrarObjetoBucket xDic, "FECHADESDE", xdesdefecha
'        RegistrarObjetoBucket xDic, "FECHAHASTA", xhastafecha
'        RegistrarObjetoBucket xDic, "LEGAJO", Self.Id
'        Call ExecuteBOSQLCUBELAYOUTDEF2( oCuboDef.LAYOUTNAME, Self.WorkSpace, xDic)
'	End If
'End Sub

REM Cubo Exportar a Calipso Analyzer - Analyzer Cubo de Liquidaciones x Legajo ( Empleado / Grupo )
Sub main()
	Stop
'	Set oCuboDef = ExisteBO(Self, "BOSQLCUBELAYOUTDEF", "LAYOUTNAME", "Analyzer Cubo de Liquidaciones x Legajo", nil, True, False, "=")
	Set oCuboDef = ExisteBO(Self, "BOSQLCUBELAYOUTDEF", "LAYOUTNAME", "Analyzer Cubo de Liquidaciones Finales", nil, True, False, "=")
    Set xDic = NewDic
    ' -------------------------------------------------------------------- '
	Set xViewGE  = NewCompoundView( Self, "GRUPOEMPLEADOS", Self.Workspace, Nil, True )
	xViewGE.addOrderColumn(NewOrderSpec(NewColumnSpec("GRUPOEMPLEADOS", "NOMBRE", ""), True ))
	xViewGE.ColumnFromPath( "NOMBRE" )
	' -------------------------------------------------------------------- '
	xFechaDesde = dateserial(year(now()), month(now()), 1 )		  ' dateserial(year(now()), month(now)-1, 1 ))
    xFechaHasta = dateserial(year(now()), month(now())+1, 1 )-1	  ' dateserial(year(now()), month(now), 1 )-1)
    ' -------------------------------------------------------------------- '
	Set xVisualvar = VisualVarEditor( "Detalle de Liquidaciones" )
	Call AddVarDate( xvisualvar, "01_dfecha"  , "Desde Fecha"      , "1 - Periodo", xFechaDesde) 'Date())   ' dateserial(year(now()), month(now), 1 ))		  ' dateserial(year(now()), month(now)-1, 1 ))
	Call AddVarDate( xvisualvar, "02_hfecha"  , "Hasta Fecha"      , "1 - Periodo", xFechaHasta) 'Date())   ' dateserial(year(now()), month(now)+1, 1 )-1)	  ' dateserial(year(now()), month(now), 1 )-1)
	Call AddVarString( xVisualVar, "08_Legajo", "Nro. Legajo", "Empleado"      , "")
	Call AddVarView( xVisualVar, "09_Grupo"   , "Grupo"      , "Grupo Empleado", xViewGE, "NOMBRE")
    
	Set oEmpresa  =  ExisteBO(Self, "COMPANIA", "ID", "{63E432B5-726E-4195-8F9D-2025C20D7467}", NIL, true, false, "=")
	xParamP = 0 : xParamLT = 0
	If "1.1" = cDbl("1.1") Then
		xParamP  = Replace(Parametro(oEmpresa, "PRESENTISMOUOM"), ",", ".")
		xParamLT = Replace(Parametro(oEmpresa, "ADICIONALLIMPTECUOM"), ",", ".")
	Else
		xParamP  = Replace(Parametro(oEmpresa, "PRESENTISMOUOM"), ".", ",")
		xParamLT = Replace(Parametro(oEmpresa, "ADICIONALLIMPTECUOM"), ".", ",")
	End If
	xParamP  = xParamP / 100
	xParamLT = xParamLT / 100
	
	xAcept = ShowVisualVar( xvisualvar )
	If xAcept Then
		xdesdefecha   = cdate(GetValueVisualVar( xvisualvar, "01_dfecha", "1 - Periodo"))
        xhastafecha   = cdate(GetValueVisualVar( xvisualvar, "02_hfecha", "1 - Periodo"))
        strdesdefecha = year(xdesdefecha) & right("00" & month(xdesdefecha),2) & right("00" & day(xdesdefecha),2) 
        strhastafecha = year(xhastafecha) & right("00" & month(xhastafecha),2) & right("00" & day(xhastafecha),2)
        xLegajo       = GetValueVisualVar(xVisualVar, "08_Legajo", "Empleado")
		xGrupo        = GetValueVisualVar(xVisualVar, "09_Grupo", "Grupo Empleado")
        xAnio         = Cstr(Year(xdesdefecha))
		xMes          = Month(xdesdefecha)
		xAnioMes      = cStr(Year(xdesdefecha)) & Right("00" & cStr(Month(xdesdefecha)), 2)
        RegistrarObjetoBucket xDic, "FECHADESDE", xdesdefecha
        RegistrarObjetoBucket xDic, "FECHAHASTA", xhastafecha
'       RegistrarObjetoBucket xDic, "AnioMes", xAnioMes
'       RegistrarObjetoBucket xDic, "Mes", xMes
		' -------------------------------------------------------------------- '
        If Not IsEmpty(xLegajo) And xLegajo <> "" Then
'		    Set oCuboDef = ExisteBO(Self, "BOSQLCUBELAYOUTDEF", "LAYOUTNAME", "Analyzer Cubo de Liquidaciones x Legajo", nil, True, False, "=")
			Set xViewEmp = NewCompoundView(Self,"EMPLEADO",Self.Workspace,Nil,True)
            xViewEmp.AddFilter(NewFilterSpec(xViewEmp.ColumnFromPath("CODIGO"), " = ", xLegajo))
			xViewEmp.AddFilter(NewFilterSpec(xViewEmp.ColumnFromPath("ACTIVESTATUS"), " <> ", 2))
            If Not xViewEmp.ViewItems.IsEmpty Then
			   For Each EE In xViewEmp.ViewItems
			   		Set oEmpleado = EE.BO
			   Next
			End If
            ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' '
            If Not oEmpleado Is Nothing Then
                RegistrarObjetoBucket xDic, "LEGAJO", oEmpleado.Id
            Else
                RegistrarObjetoBucket xDic, "LEGAJO", "00000000-0000-0000-0000-000000000000"
            End If
'			Call ExecuteBOSQLCUBELAYOUTDEF2( oCuboDef.LAYOUTNAME, Self.WorkSpace, xDic)
'			Exit Sub
		Else
			RegistrarObjetoBucket xDic, "LEGAJO", "00000000-0000-0000-0000-000000000000"
        End If
		' -------------------------------------------------------------------- '
        If Not IsEmpty(xGrupo) Then
'		    Set oCuboDef = ExisteBO(Self, "BOSQLCUBELAYOUTDEF", "LAYOUTNAME", "Cubo Mejor Remuneracion por Grupo de Empleados", nil, True, False, "=")
            Set oGrupo   = ExisteBo(Self, "GRUPOEMPLEADOS", "id", xGrupo, nil, true, false, "=")
            If Not oGrupo Is Nothing Then
                RegistrarObjetoBucket xDic, "GRUPO", oGrupo.Id
            Else
                RegistrarObjetoBucket xDic, "GRUPO", "00000000-0000-0000-0000-000000000000"
            End If
'			Call ExecuteBOSQLCUBELAYOUTDEF2( oCuboDef.LAYOUTNAME, Self.WorkSpace, xDic)
'        	Exit Sub
		Else
			RegistrarObjetoBucket xDic, "GRUPO", "00000000-0000-0000-0000-000000000000"
        End If
		' -------------------------------------------------------------------- '
		RegistrarObjetoBucket xDic, "PRESENTISMO", xParamP
        RegistrarObjetoBucket xDic, "LIMPIEZATECNICA", xParamLT
		' -------------------------------------------------------------------- '
		Call ExecuteBOSQLCUBELAYOUTDEF2( oCuboDef.LAYOUTNAME, Self.WorkSpace, xDic)
	End If
End Sub




Sub MainHide 

EXECCONTROL.Value = 1 

end sub 


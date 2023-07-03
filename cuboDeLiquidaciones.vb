' Cubo Exportar a Calipso Analyzer - Analyzer Cubo de Liquidaciones x Convenio
Sub main()
	Stop
    Set oCuboDef = ExisteBO(Self, "BOSQLCUBELAYOUTDEF", "LAYOUTNAME", "Analyzer Cubo de Liquidaciones Nuevo", nil, True, False, "=")
    Set xDic = NewDic
    ' -------------------------------------------------------------------- '
	Set xViewCN  = NewCompoundView(Self,"PERFIL",Self.Workspace,Nil,True)
	xViewCN.ColumnFromPath( "DESCRIPCION" )
	xViewCN.addOrderColumn(NewOrderSpec(NewColumnSpec("PERFIL", "DESCRIPCION", ""), True ))
	' -------------------------------------------------------------------- '
    Set xViewCC  = NewCompoundView(Self,"CENTROCOSTOS",Self.Workspace,Nil,True)
	xViewCC.ColumnFromPath( "NOMBRE" )
	xViewCC.AddFilter(NewFilterSpec(xViewCC.ColumnFromPath("ACTIVESTATUS"), " = ", 0))
	xViewCC.addOrderColumn(NewOrderSpec(NewColumnSpec("CENTROCOSTOS", "NOMBRE", ""), True ))
    ' -------------------------------------------------------------------- '
    Set xViewGE  = NewCompoundView( Self, "GRUPOEMPLEADOS", Self.Workspace, Nil, True )
	xViewGE.addOrderColumn(NewOrderSpec(NewColumnSpec("GRUPOEMPLEADOS", "NOMBRE", ""), True ))
	xViewGE.ColumnFromPath( "NOMBRE" )
    ' -------------------------------------------------------------------- '
	Set xViewTL  = NewCompoundView( Self, "TIPOLIQUIDACION", Self.Workspace, Nil, True )
	xViewTL.addOrderColumn(NewOrderSpec(NewColumnSpec("TIPOLIQUIDACION", "NOMBRE", ""), True ))
	xViewTL.ColumnFromPath( "NOMBRE" )
    ' -------------------------------------------------------------------- '
    Set xViewTP  = NewCompoundView(Self,"ITEMTIPOCLASIFICADOR",Self.Workspace,Nil,True)
    xViewTP.ColumnFromPath( "NOMBRE" )
    xViewTP.AddFilter(NewFilterSpec(xViewTP.ColumnFromPath("BO_PLACE"), " = ","{63522C61-B002-4A57-85E4-218982A6094A}"))
    xViewTP.AddFilter(NewFilterSpec(xViewTP.ColumnFromPath("ACTIVESTATUS"), " = ", 0)) 
    
    ' -------------------------------------------------------------------- '
	Set xVisualvar = VisualVarEditor( "Detalle de Liquidaciones" )
	Call AddVarDate( xvisualvar, "01_dfecha"  , "Desde Fecha"         , "A - Periodo de Tiempo"   , dateserial(cint(left(self.codigo, 4)), cint(right(self.codigo, 2)), 1 ))          ' dateserial(year(now()), month(now), 1 ))		  ' dateserial(year(now()), month(now)-1, 1 ))
	Call AddVarDate( xvisualvar, "02_hfecha"  , "Hasta Fecha"         , "A - Periodo de Tiempo"   , dateserial(cint(left(self.codigo, 4)), cint(right(self.codigo, 2)) + 1, 1 ) - 1)   ' dateserial(year(now()), month(now)+1, 1 )-1)	  ' dateserial(year(now()), month(now), 1 )-1)
	Call AddVarView( xVisualvar, "06_CN"      , "Convenio"            , "B - Convenio"            , xViewCN, "DESCRIPCION" )
	Call AddVarView( xVisualvar, "07_CC"      , "Centro de Costos"    , "C - Centro de Costos"    , xViewCC, "NOMBRE" ) 
	Call AddVarView( xVisualvar, "07_TL"      , "Tipo de Liquidación" , "D - Tipo de Liquidacion" , xViewTL, "NOMBRE" )  
    Call AddVarView( xVisualvar, "10_TP"      ,"Mensual o Jornal"     , "F - Mensual o Jornal"    , xViewTP, "NOMBRE" ) 
'   Call AddVarView( xVisualVar, "09_Grupo"   , "Grupo"               , "D - Grupo Empleado"      , xViewGE, "NOMBRE")
    Call AddVarString (xVisualVar, "08_Legajo", "Nro. Legajo"         , "E - Empleado"            , "")
    
	xAcept = ShowVisualVar( xvisualvar )
	If xAcept Then
        xdesdefecha      = cdate(GetValueVisualVar( xvisualvar, "01_dfecha", "A - Periodo de Tiempo"))
        xhastafecha      = cdate(GetValueVisualVar( xvisualvar, "02_hfecha", "A - Periodo de Tiempo"))
        strdesdefecha    = year(xdesdefecha) & right("00" & month(xdesdefecha),2) & right("00" & day(xdesdefecha),2) 
        strhastafecha    = year(xhastafecha) & right("00" & month(xhastafecha),2) & right("00" & day(xhastafecha),2)
        xConvenio        = GetValueVisualVar(xVisualVar, "06_CN", "B - Convenio")
        xCentroCosto     = GetValueVisualVar(xVisualVar, "07_CC", "C - Centro de Costos")
		xTipoLiquidacion = GetValueVisualVar(xVisualVar, "07_TL", "D - Tipo de Liquidacion")
        xTipoPago   = GetValueVisualVar(xVisualVar, "10_TP", "F - Mensual o Jornal")
'       xGrupo           = GetValueVisualVar(xVisualVar, "09_Grupo", "D - Grupo Empleado")
        xLegajo          = GetValueVisualVar(xVisualVar, "08_Legajo", "E - Empleado")
        RegistrarObjetoBucket xDic, "FECHADESDE", xdesdefecha
        RegistrarObjetoBucket xDic, "FECHAHASTA", xhastafecha
	   ' -------------------------------------------------------------------- '
        If Not IsEmpty(xConvenio) Then
            Set oConvenio = ExisteBo(Self, "PERFIL", "id", xConvenio, nil, true, false, "=")
            If Not oConvenio Is Nothing Then
                RegistrarObjetoBucket xDic, "CONVENIO", oConvenio.Id
            Else
                RegistrarObjetoBucket xDic, "CONVENIO", "00000000-0000-0000-0000-000000000000"
            End If
        Else
            RegistrarObjetoBucket xDic, "CONVENIO", "00000000-0000-0000-0000-000000000000"
        End If
	   
        If Not IsEmpty(xCentroCosto) Then
            Set oCentroCosto = ExisteBo(Self, "CENTROCOSTOS", "id", xCentroCosto, nil, true, false, "=")
            If Not oCentroCosto Is Nothing Then
                RegistrarObjetoBucket xDic, "CENTROCOSTO", oCentroCosto.Id
            Else
                RegistrarObjetoBucket xDic, "CENTROCOSTO", "00000000-0000-0000-0000-000000000000"
            End If
        Else
            RegistrarObjetoBucket xDic, "CENTROCOSTO", "00000000-0000-0000-0000-000000000000"
        End If

		If Not IsEmpty(xTipoLiquidacion) Then
            Set oTipoLiquidacion = ExisteBo(Self, "TIPOLIQUIDACION", "id", xTipoLiquidacion, nil, true, false, "=")
            If Not oTipoLiquidacion Is Nothing Then
                RegistrarObjetoBucket xDic, "TIPOLIQUIDACION", oTipoLiquidacion.Id
            Else
                RegistrarObjetoBucket xDic, "TIPOLIQUIDACION", "00000000-0000-0000-0000-000000000000"
            End If
        Else
            RegistrarObjetoBucket xDic, "TIPOLIQUIDACION", "00000000-0000-0000-0000-000000000000"
        End If
'        If Not IsEmpty(xGrupo) Then
'            Set oGrupo = ExisteBo(Self, "GRUPOEMPLEADOS", "id", xGrupo, nil, true, false, "=")
'            If Not oGrupo Is Nothing Then
'                RegistrarObjetoBucket xDic, "GRUPO", oGrupo.Id
'            Else
'                RegistrarObjetoBucket xDic, "GRUPO", "00000000-0000-0000-0000-000000000000"
'            End If
'        Else
'            RegistrarObjetoBucket xDic, "GRUPO", "00000000-0000-0000-0000-000000000000"
'        End If
        If Not IsEmpty(xTipoPago) Then      
            Set oTipoPago = Nothing
            Set oTipoPago = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "id", xTipoPago, nil, true, false, "=")
        
            If Not oTipoPago Is Nothing Then
                RegistrarObjetoBucket xDic, "MENSUALOJORNAL", oTipoPago.Id
            Else
                RegistrarObjetoBucket xDic, "MENSUALOJORNAL", "00000000-0000-0000-0000-000000000000"
            End If
        Else
           RegistrarObjetoBucket xDic, "MENSUALOJORNAL", "00000000-0000-0000-0000-000000000000"
        End If

        If Not IsEmpty(xLegajo) And xLegajo <> "" Then
		    ' -------------------------------------------------------------------- '
			Set xViewEmp  = NewCompoundView(Self,"EMPLEADO",Self.Workspace,Nil,True)
            xViewEmp.AddFilter(NewFilterSpec(xViewEmp.ColumnFromPath("CODIGO"), " = ", xLegajo))
			xViewEmp.AddFilter(NewFilterSpec(xViewEmp.ColumnFromPath("ACTIVESTATUS"), " <> ", 2))
            If Not xViewEmp.ViewItems.IsEmpty Then
			   For Each EE In xViewEmp.ViewItems
			   		Set oEmpleado = EE.BO
			   Next
			End If
            ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' '
'            Set oEmpleado = ExisteBo(Self, "EMPLEADO", "CODIGO", xLegajo, nil, true, false, "=")
            If Not oEmpleado Is Nothing Then
                RegistrarObjetoBucket xDic, "LEGAJO", oEmpleado.Id
            Else
                RegistrarObjetoBucket xDic, "LEGAJO", "00000000-0000-0000-0000-000000000000"
            End If
        Else
            RegistrarObjetoBucket xDic, "LEGAJO", "00000000-0000-0000-0000-000000000000"
        End If

        Call ExecuteBOSQLCUBELAYOUTDEF2( oCuboDef.LAYOUTNAME, Self.WorkSpace, xDic)
	   'Call ExecuteBOSQLCUBELAYOUTDEFParam( oCuboDef.LAYOUTNAME, Self.WorkSpace, xDic, "c:\CuboUsuarioCF" )
	End If
End Sub

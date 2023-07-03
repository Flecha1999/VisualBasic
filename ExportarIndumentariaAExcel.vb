SUB MAIN
	STOP

	xstring = StringConexion( "CALIPSO", self.WorkSpace ) 
	set xcone = createobject("adodb.connection")
	xcone.connectionstring  = xstring
	xcone.connectiontimeout = 0

	set xview = NewCompoundView( self, "Centrocostos", self.Workspace, nil, true )
	xView.addfilter( NewFilterSpec(NewColumnSpec( "Centrocostos", "activestatus", "Centrocostos" )," = ", "0"))
	XVIeW.ADDORDERCOLUMN(NewOrderSpec( NewColumnSpec( "Centrocostos", "codigo", "Centrocostos" ), FALSE )) 
	set xColumnacodigo = xView.ColumnFromPath("codigo")
	set xColumnanombre = xView.ColumnFromPath("Nombre")

	set xVisualVar = VisualVarEditor( "Seleccion el Tipo de Producto" )
    
	
	call AddVarView   ( xVisualVar, "CentroCostos", "CentroCostos","CentroCostos" , xView, "codigo%nombre")
 	
	XAcept = ShowVisualVar( xVisualVar )

	if xacept then
	   cc  = GetValueVisualVar( xVisualVar, "CentroCostos", "CentroCostos" )
	   if not isempty(cc) then
	   	  filtro = " where vp.centrocostos_id='"&cc&"' "
	   else
	      filtro = ""
	   end if
    else
	   call msgbox("Proceso Cancelado por el usuario")
	   exit sub
    End IF

	
	set xRs = RecordSet(xCone, "select top 1 * from producto")
	xRs.close
	xRs.activeconnection.commandtimeout=0
	
	xquery =    "SELECT CC.NOMBRE AS CENTROCOSTOS,EMPL.CODIGO AS LEGAJO,PER.NOMBRE AS EMPLEADO,"&chr(13)&_
                "ISNULL(PUE.DESCRIPCION, '') AS PUESTO,"&chr(13)&_
                "ISNULL(UEMPL.OBJTALLEBUZO_N,'') AS TalleBuzo,"&chr(13)&_
                "ISNULL(UEMPL.OBJTALLECALZADO_N,'') AS TalleCalzado,"&chr(13)&_
                "ISNULL(UEMPL.OBJTALLECAMISA_N,'') AS TalleCamisa,"&chr(13)&_
                "ISNULL(UEMPL.OBJTALLECAMPERA_N,'') AS TalleCampera,"&chr(13)&_
                "ISNULL(UEMPL.OBJTALLEPANTALON_N,'')AS TallePantalon,"&chr(13)&_
                "ISNULL(UEMPL.OBJTALLEPRENDACOMPLETA_N,'') AS TallePrendaCompleta,"&chr(13)&_
                "ISNULL(UEMPL.OBJTALLEREMERA_N,'') AS TalleRemera"&chr(13)&_
                "FROM EMPLEADO AS EMPL WITH(NOLOCK)"&chr(13)&_
                "INNER JOIN UD_EMPLEADO AS UEMPL WITH(NOLOCK) ON EMPL.BOEXTENSION_ID = UEMPL.ID"&chr(13)&_
                "INNER JOIN PERSONAFISICA AS PER WITH(NOLOCK) ON EMPL.ENTEASOCIADO_ID = PER.ID"&chr(13)&_
                "INNER JOIN CENTROCOSTOS AS CC WITH(NOLOCK) ON EMPL.CENTROCOSTOS_ID = CC.ID"&chr(13)&_
                "LEFT JOIN SECTOR AS SEC WITH(NOLOCK) ON EMPL.SECTOR_ID = SEC.ID"&chr(13)&_
                "LEFT JOIN PUESTO AS PUE WITH(NOLOCK) ON EMPL.PUESTO_ID = PUE.ID"&chr(13)&_
                "WHERE EMPL.ACTIVESTATUS = 0 "&chr(13)&_
                "ORDER BY CC.NOMBRE, PER.NOMBRE, EMPL.CODIGO"
    
	xRs.source = xquery
	xRs.open

	call ProgressControl(Self.Workspace, "Informe Antiguedad" , 0, 300)
	Set HojaExcel = CreateObject("Excel.Application")
	HojaExcel.Workbooks.Add
	
	'------HOJA 1---- QUERY1
	HojaExcel.Sheets("Hoja1").Select
	HojaExcel.ActiveSheet.Cells(1, 1).Value = "ISCOT" 
    With HojaExcel.ActiveSheet.Range("C1:Y1")
        .Merge
        .Value = "Relevamiento de indumentaria"
    End With
    
    ' Combina las celdas desde la 2C hasta la 2E
    With HojaExcel.ActiveSheet.Range("C2:E2")
        .Merge
        .Value = "Camisa/Chaqueta"
    End With	
    With HojaExcel.ActiveSheet.Range("F2:H2")
        .Merge
        .Value = "Pantaón"
    End With	
    With HojaExcel.ActiveSheet.Range("I2:K2")
        .Merge
        .Value = "Calzado"
    End With	
    With HojaExcel.ActiveSheet.Range("L2:N2")
        .Merge
        .Value = "Campera"
    End With	 
    With HojaExcel.ActiveSheet.Range("O2:Q2")
        .Merge
        .Value = "Buzo/Sweater"
    End With	
    With HojaExcel.ActiveSheet.Range("R2:T2")
        .Merge
        .Value = "Remera/Chomba"
    End With
    With HojaExcel.ActiveSheet.Range("U2:W2")
        .Merge
        .Value = "Prenda Completa"
    End With
    HojaExcel.ActiveSheet.Cells(2, 24).Value = "Cabeza"
    HojaExcel.ActiveSheet.Cells(2, 25).Value = "Bolzo"

	HojaExcel.ActiveSheet.Cells(3, 3).Value = "C"
    HojaExcel.ActiveSheet.Cells(3, 4).Value = "T"
    HojaExcel.ActiveSheet.Cells(3, 5).Value = "N"

    HojaExcel.ActiveSheet.Cells(3, 6).Value = "C"
    HojaExcel.ActiveSheet.Cells(3, 7).Value = "T"
    HojaExcel.ActiveSheet.Cells(3, 8).Value = "N"

    HojaExcel.ActiveSheet.Cells(3, 9).Value = "C"
    HojaExcel.ActiveSheet.Cells(3, 10).Value = "T"
    HojaExcel.ActiveSheet.Cells(3, 11).Value = "N"

    HojaExcel.ActiveSheet.Cells(3, 12).Value = "C"
    HojaExcel.ActiveSheet.Cells(3, 13).Value = "T"
    HojaExcel.ActiveSheet.Cells(3, 14).Value = "N"

    HojaExcel.ActiveSheet.Cells(3, 15).Value = "C"
    HojaExcel.ActiveSheet.Cells(3, 16).Value = "T"
    HojaExcel.ActiveSheet.Cells(3, 17).Value = "N"

    HojaExcel.ActiveSheet.Cells(3, 18).Value = "C"
    HojaExcel.ActiveSheet.Cells(3, 19).Value = "T"
    HojaExcel.ActiveSheet.Cells(3, 20).Value = "N"

    HojaExcel.ActiveSheet.Cells(3, 21).Value = "C"
    HojaExcel.ActiveSheet.Cells(3, 22).Value = "T"
    HojaExcel.ActiveSheet.Cells(3, 23).Value = "N"

    HojaExcel.ActiveSheet.Cells(3, 24).Value = "C"
    HojaExcel.ActiveSheet.Cells(3, 25).Value = "C"
	
	R = 4
	
	do while not xRs.eof
	    call ProgressControlAvance(Self.Workspace, "Empleado: " & CStr(xRs("Legajo").Value)&" "& CStr(xRs("EMPLEADO").Value))
		
		HojaExcel.ActiveSheet.Cells(R, 1).Value = CStr(xRs("Legajo").Value)
		HojaExcel.ActiveSheet.Cells(R, 2).Value = CStr(xRs("EMPLEADO").Value)
		HojaExcel.ActiveSheet.Cells(R, 4).Value = CStr(xRs("TalleCamisa").Value)
		HojaExcel.ActiveSheet.Cells(R, 7).Value = CStr(xRs("TallePantalon").Value)
		HojaExcel.ActiveSheet.Cells(R, 10).Value = CStr(xRs("TalleCalzado").Value) '
		HojaExcel.ActiveSheet.Cells(R, 13).Value = CStr(xRs("TalleCampera").Value) '"Expediente"
		HojaExcel.ActiveSheet.Cells(R, 16).Value = CStr(xRs("TalleBuzo").Value) '"Total"
        HojaExcel.ActiveSheet.Cells(R, 19).Value = CStr(xRs("TalleRemera").Value) '
		HojaExcel.ActiveSheet.Cells(R, 22).Value = CStr(xRs("TallePrendaCompleta").Value) '"Expediente"
		R = R + 1
			  
		xRs.MOVENEXT
	loop
	'HojaExcel.ActiveSheet.Columns("A:L").AutoFit
	call ProgressControlFinish(Self.Workspace)
	HojaExcel.Visible 	= true
	set HojaExcel 		= nothing
end sub

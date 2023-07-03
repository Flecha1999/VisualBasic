REM ISC - TXT SIJP (GENERAL)
Sub Main
	Stop
    Set xApp	= Self : xTotal = 0
	Set fso 	= CreateObject("Scripting.FileSystemObject")
	Path 		= "C:\Util\TXT SIJP\"
	NOMBRE	    =	"SIJP_" & Mid(Date(), 1, 2) & Mid(Date(), 4, 2) & Mid(Date(), 7, 4) &  NombreUsuario()&".TXT"	' &cstr(hour(time))&cstr(minute(time))
	xNombreTxt	=	"SIJP_" & Mid(Date(), 1, 2) & Mid(Date(), 4, 2) & Mid(Date(), 7, 4) & NombreUsuario()&".txt"
	xNombreXls	=	"SIJP_" & Mid(Date(), 1, 2) & Mid(Date(), 4, 2) & Mid(Date(), 7, 4) &  NombreUsuario()&".xls"
	' ---------------------------------------------------- '
	Set xViewTL	= NewCompoundView(Self, "tipoliquidacion", Self.Workspace, nil, True)
	Set xColum	= xViewTL.ColumnFromPath( "Nombre" )
	' ---------------------------------------------------- '
	Set xVisualVar = VisualVarEditor( "TXT SIJP" )
'	Call AddVarDate  ( xvisualvar, "01_dfecha" , "Desde Fecha"       , "Parametros", dateserial(year(now()),month(now)-1,1 ))
'	Call AddVarDate  ( xvisualvar, "02_hfecha" , "Hasta Fecha"        , "Parametros", dateserial(year(now()),month(now),1 )-1)
	Call AddVarString( xvisualvar, "03_Path"    , "Path"                    , "Parametros", Path )
'	Call AddVarView  ( xvisualvar, "04_TipoLiq", "Tipo Liquidacion"  , "Parametros" , xViewTL, "NOMBRE" )
	xAcept = ShowVisualVar( xvisualvar )
	If xAcept Then
'		xdesdefecha			= Liquidacion.InicioPeriodo	' GetValueVisualVar( xvisualvar, "01_dfecha", "Parametros" )
'		strdesdefecha		= Year(cDate(xdesdefecha)) & right("00" & month(cDate(xdesdefecha)),2) & right("00" & day(cDate(xdesdefecha)),2)
'		xhastafecha			= Liquidacion.FinPeriodo	' GetValueVisualVar( xvisualvar, "02_hfecha", "Parametros" )
'		strhastafecha		= Year(cDate(xdesdefecha)) & right("00" & month(cDate(xhastafecha)),2) & right("00" & day(cDate(xhastafecha)),2)
		xPath					= GetValueVisualVar( xvisualvar, "03_Path", "Parametros" )
'		xTipoLiquidacion	= GetValueVisualVar( xVisualVar, "04_TipoLiq", "Parametros" )
		Path = xPath
		If (fso.FileExists(path & xNombreTxt)) Then
			xrta = msgbox("El archivo ya existe. ¿Sobreescribe?",4)
			If xrta = 6 Then 'dijo si
				fso.DeleteFile(path & xNombreTxt)
				If (fso.FileExists(path & xNombreXls)) Then
					fso.DeleteFile(path & xNombreXls)
				End If
			Else
				msgbox "Proceso abortado"
				Exit Sub
			End If
		End If
		Stop
		xstringCON = stringconexion("CALIPSO",xapp.workspace)
		Set xcone = createobject("adodb.connection")
		xcone.connectionstring = xstringCON
		xcone.open
		Set xrst = recordSet(xstringCON,"select top 1 ID from producto ")
		xrst.close
		If UCase(NombreUsuario()) = "DESARROLLO2" Then
		    xFiltroAux = " AND E.CODIGO IN ( '7802' ) "   '" AND E.CODIGO IN ( '867', '10134', '7684', '8201', '8276', '9749', '10091', '6231', '10118', '10187', '10195' ) "
		Else
			xFiltroAux = ""
		End If
'									   " 			case when CONC.CODIGO in( '9853' ) Then SUM(dliq.CANTIDAD_CANTIDAD) Else 0 END Cant_HorasExtras, "&_
		xquery =  " SELECT Q.IDEMPLEADO AS 'IDEMPLEADO', "&_
						" Q.LEGAJO AS 'LEGAJO',  "&_
						" Q.CUIT AS 'CUIT',  "&_
						" Q.NOMBREEMPLEADO AS 'NOMBRE_EMPLEADO',  "&_
						" Q.CONYUGE AS 'CONYUGE',  "&_
						" Q.SITUACION AS 'SITUACION',  "&_
						" Q.CONDICION AS 'CONDICION',  "&_
						" '49' AS 'CODIGO_ACTIVIDAD', "&_
						" Q.CODIGOZONA AS 'CODIGO_ZONA',  "&_
						" Q.PORCENTAJE_APORTE_OBRASOCIAL AS 'PORCENTAJE_APORTE_OBRASOCIAL', "&_
						" Q.CODIGOMODALIDADCONTRATACION AS 'CODIGO_MODALIDAD_CONTRATACION',  "&_
						" Q.CODIGOOBRASOCIAL AS 'CODIGO_OBRA_SOCIAL',  "&_
						" Q.ASIGFAMILIRPAGADA AS 'ASIGNACION_FAMILIAR_PAGA',  "&_
						" Q.IAPORTEVOLUNTARIO AS 'IMPORTE_APORTE_VOLUNTARIO',  "&_
						" Q.IADICIONALOS AS 'IMPORTE_ADICIONAL_OBRA_SOCIAL',  "&_
						" Q.IEXCAPORTEOS AS 'IMPORTE_EXC_APORTE_OS',  "&_
						" SUM(Q.IEXCAPORTESS) AS 'IMPORTE_APORTE_SS',  "&_
						" Q.LOCALIDAD AS 'LOCALIDAD',  "&_
						" Q.CODIGOSINIESTRO AS 'CODIGO_SINIESTRO',  "&_
						" Q.CORRESPONDEREDUCCION AS 'CORRESPONDE_REDUCCION',  "&_
						" Q.CAPITALRECOMPOSICION AS 'CAPITAL_RECOMPOSICION',  "&_
						" Q.TIPOEMPRESA AS 'TIPO_EMPRESA',  "&_
						" Q.ADICIONALOS AS 'ADICIONAL_OS', "&_
						" Q.REGIMEN AS 'REGIMEN',  "&_
						" Q.SITUACION1 AS 'SITUACION_1', "&_
						" Q.SITUACION2 AS 'SITUACION_2',  "&_
						" Q.SITUACION3 AS 'SITUACION_3',  "&_
						" Q.DIASITREVISTA1 AS 'DIA_REVISTA_1',  "&_
						" Q.DIASITREVISTA2 AS 'DIA_REVISTA_2' ,  "&_
						" Q.DIASITREVISTA3 AS 'DIA_REVISTA_3',  "&_
						" Q.CONVENCIONADO AS 'CONVENCIONADO',  "&_
						" Q.REMUNERACION6 AS 'REMUNERACION_6',  "&_
						" Q.TIPOOPERACION AS 'TIPO_OPERACION',  "&_
						" Q.REMUNERACION7 AS 'REMUNERACION_7',  "&_
						" SUM(Q.Rem_Maternidad) AS 'MATERNIDAD', "&_ 
						" Q.RECTIFICACION AS 'RECTIFICACION', "&_
						" Q.CONTRIBUCIONTAREADIf AS 'CONTRIBUCION_TAREA_DIF',  "&_ 
						" CAST(SUM(Q.HORASTRABAJADAS) AS INTEGER) AS 'HORAS_TRABAJADAS', "&_ 
						" Q.SEGUROVIDA AS 'SEGURO_VIDA', "&_
						" SUM(Q.REM_1) AS 'REM_1',  "&_
						" SUM(Q.REM_2) AS 'REM_2',   "&_	
						" SUM(Q.REM_3) AS 'REM_3',  "&_
						" SUM(Q.REM_4) AS 'REM_4',  "&_
						" SUM(Q.REM_5) AS 'REM_5',   "&_
						" SUM(Q.REM_8) AS 'REM_8',  	 "&_
						" SUM(Q.REM_9) AS 'REM_9',   "&_	
						" SUM(Q.REM_TOTAL) AS 'REMUNERACION_TOTAL', "&_
						" SUM(Q.SAC) AS 'SAC', "&_
						" SUM(Q.HorasExtras) AS 'HORAS_EXTRAS', "&_
						" SUM(Q.Vacaciones) AS 'VACACIONES',  "&_
						" SUM(Q.DiasNotrabajados) AS 'DIAS_TRABAJADOS',  "&_
						" SUM(Q.Sueldo) AS 'SUELDO',  "&_
						" SUM(Q.Adicional) AS 'ADICIONAL',  "&_
						" SUM(Q.Premio) AS 'PREMIO',    "&_
						" SUM(Q.Cant_HorasExtras) AS 'CANT_HORAS_EXTRAS', "&_
						" SUM(Q.No_Remunerativo) AS 'NO_REMUNERATIVO', "&_	
						" '000000000,00' AS 'ZONA_DESFAVORABLE', "&_
						" SUM(Q.Importe_Detraer) AS Importe_Detraer "&_
						" FROM (  "&_
							   " Select E.ID AS IDEMPLEADO, e.codigo AS LEGAJO, pf.cuit AS CUIT, "&_
							   " 			PF.NOMBRE as NOMBREEMPLEADO,  "&_
							   " 			case es.estadocivil when 'Casado/a' Then '1' Else '0' end  as CONYUGE, "&_
							   " 			isnull(sit.CODIGO,0) as SITUACION, "&_
							   " 			isnull(COND.CODIGO,0) as CONDICION, "&_
							   " 			ISNULL(Z.CODIGO, '18') AS CODIGOZONA, "&_
							   " 			'00,00' AS PORCENTAJE_APORTE_OBRASOCIAL, "&_
							   " 			ISNULL(TCONT.CODIGO, '000') AS CODIGOMODALIDADCONTRATACION, "&_
							   " 			ISNULL(OS.CODIGO, '000000') AS CODIGOOBRASOCIAL, "&_
							   " 			'000000,00' AS ASIGFAMILIRPAGADA, "&_
							   " 			'000000,00' AS IAPORTEVOLUNTARIO, "&_
							   " 			'000000,00' AS IADICIONALOS, "&_
							   "			case when CONC.CODIGO in ( '378', '1637', '744', '1184', '1486', '2172' ) Then SUM(ABS(dliq.VALOR_IMPORTE)) Else 0 END AS IEXCAPORTESS, "&_
							   " 			'000000,00' AS IEXCAPORTEOS, "&_
							   " 			ISNULL(Z.CODIGO, '') AS LOCALIDAD, "&_
							   " 			ISNULL(SINIESTRO.CODIGO, 0) AS CODIGOSINIESTRO, "&_
							   " 			E.CORRESPONDEREDUCCION AS CORRESPONDEREDUCCION, "&_
							   " 			'000000,00' AS CAPITALRECOMPOSICION, "&_
							   " 			'4' AS TIPOEMPRESA, "&_
							   " 			case when ud.adicionalASOSECAC > 0 Then ud.adicionalASOSECAC Else 0 end as ADICIONALOS, "&_
							   " 			case when E.REGIMEN <> '' Then E.REGIMEN Else '1' end as REGIMEN, "&_
							   " 			E.SITREVISTA1 AS SITUACION1, "&_
							   " 			E.DIASITREVISTA1 AS DIASITREVISTA1, "&_
							   " 			E.SITREVISTA2 AS SITUACION2, "&_
							   " 			E.DIASITREVISTA2 AS DIASITREVISTA2, "&_
							   " 			E.SITREVISTA3 AS SITUACION3, "&_
							   " 			E.DIASITREVISTA3 AS DIASITREVISTA3, "&_
							   " 			CASE WHEN E.CONVENCIONADO = 'F' Then '0' Else '1' END AS CONVENCIONADO, "&_
							   " 			'000000000,00' AS REMUNERACION6,  "&_
							   " 			'0' AS TIPOOPERACION,  "&_
							   " 			'000000000,00' AS REMUNERACION7, "&_
							   " 			'000000,00' AS RECTIFICACION, "&_
							   " 			'000000,00' AS CONTRIBUCIONTAREADIF, "&_
							   "            case when CONC.CODIGO in( '9856' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END HORASTRABAJADAS, "&_
							   " 			'1' AS SEGUROVIDA, "&_
							   " 			case when CONC.CODIGO in( '9500' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END REM_1, "&_   
							   " 			case when CONC.CODIGO in( '9550' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END REM_2, "&_	 
							   " 			case when CONC.CODIGO in( '9600' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END REM_3,  "&_	 
							   " 			case when CONC.CODIGO in( '9650' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END REM_4,  "&_	 
							   " 			case when CONC.CODIGO in( '9700' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END REM_5,  "&_	 
							   " 			case when CONC.CODIGO in( '9750' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END REM_8,  "&_	 
							   " 			case when CONC.CODIGO in( '9800' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END REM_9, "&_ 	 
							   " 			case when CONC.CODIGO in( '9850' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END REM_TOTAL, "&_
							   " 			case when ISIJP.CODIGO in('02')  Then SUM(dliq.VALOR_IMPORTE) Else 0 END SAC, "&_
							   " 			case when ISIJP.CODIGO in('03') Then SUM(dliq.VALOR_IMPORTE) Else 0 END HorasExtras, "&_
							   " 			case when ISIJP.CODIGO in('05') Then SUM(dliq.VALOR_IMPORTE) Else 0 END Vacaciones,  "&_
							   " 			case when CONC.CODIGO in( '9855' ) Then SUM(dliq.CANTIDAD_CANTIDAD) Else 0 END DiasNotrabajados, "&_
							   " 			case when ISIJP.CODIGO in('01') Then SUM(dliq.VALOR_IMPORTE) Else 0 END Sueldo,  "&_
							   " 			case when ISIJP.CODIGO in('07') Then SUM(dliq.VALOR_IMPORTE) Else 0 END Adicional,  "&_
							   " 			case when ISIJP.CODIGO in('08') Then SUM(dliq.VALOR_IMPORTE) Else 0 END Premio,    "&_
							   " 			case when ISIJP.CODIGO in('03')    Then SUM(dliq.CANTIDAD_CANTIDAD) Else 0 END Cant_HorasExtras, "&_
							   " 			case when CONC.CODIGO in( '9859' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END No_Remunerativo, "&_
							   " 			case when CONC.CODIGO in( '9860' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END Rem_Maternidad, "&_
                               "            case when CONC.CODIGO in( '9870' ) Then SUM(dliq.VALOR_IMPORTE) Else 0 END Importe_Detraer "&_
							   "  FROM EMPLEADO E with (nolock) "&_
							   " INNER JOIN UD_EMPLEADO             UD with(nolock) ON ud.id = e.BOEXTENSION_ID "&_
							   "  LEFT JOIN SITUACION              SIT with(nolock) ON sit.id = e.situacion_ID   "&_
							   "  LEFT JOIN CONDICION             COND with(nolock) ON COND.id = E.CONDICIONEMPLEADO_ID "&_
							   "  LEFT JOIN OBRASOCIAL              OS with(nolock) ON E.OBRASOCIAL_ID = OS.ID  AND OS.ID <> '6C2F2986-394E-43C1-A9C4-69F325512389'  "&_
							   "  LEFT JOIN TIPOCONTRATACION     TCONT with(nolock) ON E.TIPOCONTRATACION_ID = TCONT.ID  "&_
							   "  LEFT JOIN ZONA                     Z with(nolock) ON z.iD = E.zona_ID "&_
							   "  LEFT JOIN SINIESTRO        SINIESTRO with(nolock) ON siniestro.id = E.siniestro_id "&_
							   " INNER JOIN PERSONAFISICA           PF with(nolock) ON pf.id = e.enteasociado_id 	 "&_
							   "  LEFT JOIN ESTADOCIVIL             ES with(nolock) ON ES.ID = PF.ESTADOCIVIL_ID   "&_ 
							   " INNER JOIN RESUMENLIQUIDACION       R with(nolock) ON e.id = R.empleado_id "&_   
							   " INNER JOIN DETALLELIQUIDACION    DLIQ with(nolock) ON R.ITEMSDETALLELIQUIDACION_ID = DLIQ.BO_PLACE_ID 	 "&_  
							   " INNER JOIN CONCEPTO              CONC with(nolock) ON DLIQ.CONCEPTO_ID = CONC.ID and ( conc.TIPOCONCEPTO_ID  <> '3AFCDD75-53B4-470E-B385-9CA352B8344A' OR CONC.CODIGO IN ( '378', '1637', '744', '1184', '1486', '2172' ) ) "&_
							   " iNNER JOIN TIPOCONCEPTO            TC with(nolock) ON TC.id = conc.TIPOCONCEPTO_ID  "&_
							   " INNER JOIN UD_CONCEPTOSUELDO   UDCONC with(nolock) ON UDCONC.ID = CONC.BOEXTENSION_ID "&_
							   "  LEFT JOIN ITEMTIPOCLASIFICADOR ISIJP with(nolock) ON isijp.ID = udconc.TIPOCONCEPTOSIJP_ID "&_
							   " INNER JOIN LIQUIDACION              L with(nolock) ON L.id = R.liquidacion_id 	    "&_
							   " INNER JOIN TIPOLIQUIDACION         LI with(nolock) ON LI.id= L.tipoliquidacion_ID  "&_	  
							   " INNER JOIN SJCARPETA               SJ with(nolock) ON SJ.LIQUIDACIONES_ID = L.BO_PLACE_ID "&_	
							   " WHERE r.estado = 0 "&_
							   "   AND sj.id = '"&Self.ID&"' " & xFiltroAux &_
							   "   AND R.NETO_IMPORTE >= -0.01 "&_
							   "   AND LI.id <> 'F1036414-183B-4299-BEAB-94162E35D561' " &_
							   " GROUP BY E.ID ,PF.NOMBRE,es.estadocivil,e.codigo, pf.cuit, CONC.CODIGO ,sit.CODIGO, COND.CODIGO, Z.CODIGO, TCONT.CODIGO, OS.CODIGO,Z.NOMBRE, "&_
							   " 			 SINIESTRO.CODIGO, E.CORRESPONDEREDUCCION, E.CAPITALRECOMPOSICION, ud.adicionalASOSECAC, E.REGIMEN, E.SITREVISTA1, E.SITREVISTA2, E.SITREVISTA3, "&_
							   " 			 E.DIASITREVISTA1, E.DIASITREVISTA2, E.DIASITREVISTA3,E.CONVENCIONADO, ISIJP.CODIGO  "&_
							   " ) AS Q 	  "&_
						" GROUP BY Q.IDEMPLEADO, Q.LEGAJO, Q.CUIT, Q.NOMBREEMPLEADO, Q.CONYUGE, Q.SITUACION, Q.CONDICION, Q.CODIGOZONA,Q.PORCENTAJE_APORTE_OBRASOCIAL , Q.CODIGOMODALIDADCONTRATACION, Q.CODIGOOBRASOCIAL, Q.ASIGFAMILIRPAGADA, Q.IAPORTEVOLUNTARIO, "&_
						" 			  Q.IADICIONALOS, Q.IEXCAPORTEOS, Q.LOCALIDAD, Q.CODIGOSINIESTRO, Q.CORRESPONDEREDUCCION, Q.CAPITALRECOMPOSICION, Q.TIPOEMPRESA, Q.ADICIONALOS,Q.REGIMEN, Q.SITUACION1, "&_
						" 			  Q.SITUACION2, Q.SITUACION3, Q.DIASITREVISTA1, Q.DIASITREVISTA2, Q.DIASITREVISTA3, Q.CONVENCIONADO, Q.REMUNERACION6, Q.TIPOOPERACION, Q.REMUNERACION7, Q.RECTIFICACION, "&_
						" 			  Q.CONTRIBUCIONTAREADIF, Q.SEGUROVIDA	 "&_
						" ORDER BY Q.CUIT 	" 
'				  	   "     AND L.INICIOPERIODO >= '"&strdesdefecha&"' 	 "&_ 
'				  	   "     AND L.FINPERIODO <= '"&strhastafecha&"' 	 "&_
'					   AND E.CODIGO IN ( '867', '10134', '7684', '8201', '8276', '9749', '10091', '6231', '10118', '10187', '10195' ) 
			xRst.activeconnection.commandtimeout = 0
			xRst.source = xquery
			xRst.open
			' *************************************************************************************************** '
			Do While Not xRst.Eof
				xTotal = xTotal + 1 : xRst.movenext
			Loop
			xRst.MoveFirst
			' *************************************************************************************************** '
			Call ProgressControl(Self.Workspace, "Exportar TXT SIJP" , 0, xTotal + 1 )
			Call ProgressControlAvance(Self.Workspace, "Procesando: Extrayendo datos!.")
			Call ProgressControlStatusText(Self.WorkSpace, "Procesando: Extrayendo datos!.")
			' *************************************************************************************************** '
			xExcel = False
			If MsgBox("Generar Excel de control?" , 36, "Pregunta") = 6 Then
				xExcel = True
				Set HojaExcel = CreateObject("Excel.Application")
				HojaExcel.Workbooks.add
				HojaExcel.Worksheets.add
				HojaExcel.Visible = False ' True
				HojaExcel.activesheet.name = "931"
				' -------------------------------------- '
				HojaExcel.ActiveSheet.Cells(1,  1).Value = "Cuit"
				HojaExcel.ActiveSheet.Cells(1,  2).Value = "Nombre"
				HojaExcel.ActiveSheet.Cells(1,  3).Value = "Conyuge"
				HojaExcel.ActiveSheet.Cells(1,  4).Value = "Can_Hijos"
				HojaExcel.ActiveSheet.Cells(1,  5).Value = "Cod_Situacion"
				HojaExcel.ActiveSheet.Cells(1,  6).Value = "Cod_Condicion"
				HojaExcel.ActiveSheet.Cells(1,  7).Value = "Cod_Actividad"
				HojaExcel.ActiveSheet.Cells(1,  8).Value = "Cod_Zona"
				HojaExcel.ActiveSheet.Cells(1,  9).Value = "% Aporte OS"
				HojaExcel.ActiveSheet.Cells(1, 10).Value = "Tipo_Contrat."
				HojaExcel.ActiveSheet.Cells(1, 11).Value = "Obra_Social"
				HojaExcel.ActiveSheet.Cells(1, 12).Value = "Cant_Adher."
				HojaExcel.ActiveSheet.Cells(1, 13).Value = "Rem_Total"
				HojaExcel.ActiveSheet.Cells(1, 14).Value = "Rem_1"
				HojaExcel.ActiveSheet.Cells(1, 15).Value = "Asig_Fam_Pag"
				HojaExcel.ActiveSheet.Cells(1, 16).Value = "Imp_Ap_Vol"
				HojaExcel.ActiveSheet.Cells(1, 17).Value = "Imp_Adic_OS"
				HojaExcel.ActiveSheet.Cells(1, 18).Value = "Imp_Ex_SS"
				HojaExcel.ActiveSheet.Cells(1, 19).Value = "Imp_Ex_OS"
				HojaExcel.ActiveSheet.Cells(1, 20).Value = "Provincia"
				HojaExcel.ActiveSheet.Cells(1, 21).Value = "Rem_2"
				HojaExcel.ActiveSheet.Cells(1, 22).Value = "Rem_3"
				HojaExcel.ActiveSheet.Cells(1, 23).Value = "Rem_4"
				HojaExcel.ActiveSheet.Cells(1, 24).Value = "Cod_Stro"
				HojaExcel.ActiveSheet.Cells(1, 25).Value = "Corresp_Reduc"
				HojaExcel.ActiveSheet.Cells(1, 26).Value = "Cap_Recomp."
				HojaExcel.ActiveSheet.Cells(1, 27).Value = "Tipo_Emp"
				HojaExcel.ActiveSheet.Cells(1, 28).Value = "Ap_Adic_OS"
				HojaExcel.ActiveSheet.Cells(1, 29).Value = "Regimen"
				HojaExcel.ActiveSheet.Cells(1, 30).Value = "Sit_Rev_1"
				HojaExcel.ActiveSheet.Cells(1, 31).Value = "Dias_ST1"
				HojaExcel.ActiveSheet.Cells(1, 32).Value = "Sit_Rev_2"
				HojaExcel.ActiveSheet.Cells(1, 33).Value = "Dias_ST2"
				HojaExcel.ActiveSheet.Cells(1, 34).Value = "Sit_Rev_3"
				HojaExcel.ActiveSheet.Cells(1, 35).Value = "Dias_ST3"
				HojaExcel.ActiveSheet.Cells(1, 36).Value = "Sueldo"
				HojaExcel.ActiveSheet.Cells(1, 37).Value = "SAC"
				HojaExcel.ActiveSheet.Cells(1, 38).Value = "Hs_Extras"
				HojaExcel.ActiveSheet.Cells(1, 39).Value = "Zona_Desf"
				HojaExcel.ActiveSheet.Cells(1, 40).Value = "Vacaciones"
				HojaExcel.ActiveSheet.Cells(1, 41).Value = "Dias_Trab"
				HojaExcel.ActiveSheet.Cells(1, 42).Value = "Rem_5"
				HojaExcel.ActiveSheet.Cells(1, 43).Value = "Convencionado"
				HojaExcel.ActiveSheet.Cells(1, 44).Value = "Rem_6"
				HojaExcel.ActiveSheet.Cells(1, 45).Value = "Tipo_Oper"
				HojaExcel.ActiveSheet.Cells(1, 46).Value = "Adicional"
				HojaExcel.ActiveSheet.Cells(1, 47).Value = "Premios"
				HojaExcel.ActiveSheet.Cells(1, 48).Value = "Rem_8"
				HojaExcel.ActiveSheet.Cells(1, 49).Value = "Rem_7"
				HojaExcel.ActiveSheet.Cells(1, 50).Value = "Cant_Hs_Ext"
				HojaExcel.ActiveSheet.Cells(1, 51).Value = "No_Remun"
				HojaExcel.ActiveSheet.Cells(1, 52).Value = "Maternidad"
				HojaExcel.ActiveSheet.Cells(1, 53).Value = "Rectificacion"
				HojaExcel.ActiveSheet.Cells(1, 54).Value = "Rem_9"
				HojaExcel.ActiveSheet.Cells(1, 55).Value = "Tarea_Dif"
				HojaExcel.ActiveSheet.Cells(1, 56).Value = "Hs_Trab"
				HojaExcel.ActiveSheet.Cells(1, 57).Value = "Seg_Colec"
                HojaExcel.ActiveSheet.Cells(1, 58).Value = "Imp_Detraer"
				HojaExcel.ActiveSheet.Cells(1, 59).Value = "Inc_Solidario"
                HojaExcel.ActiveSheet.Cells(1, 60).Value = "Rem_11"
			End If
			Fila = 2  :  xContador = 1
		    dim xcampopro(60)
		    while Not xrst.eof
						'BUSCO EL EMPLEADO
						Set oEmpleado = Nothing
						xLegajo       = xrst("LEGAJO").value
						If ( xLegajo = "8319"  ) And UCase(NombreUsuario()) = "DESARROLLO2" Then
						   MsgBox "Stop"
						   Stop
						End If
						SendDebug ("Contador " & xContador & "  --  Legajo: " & xLegajo)
						Set xView = NewCompoundView(Self,"EMPLEADO", Self.Workspace, nil, True)
						xView.addfilter( NewFilterSpec(NewColumnSpec( "EMPLEADO", "codigo", "EMPLEADO" )," = ", xLegajo))
						For Each ee In xView.viewItems
							Set oEmpleado = ee.bo
							Exit For
						Next
						xCantHijos = 0
						If Not oEmpleado Is Nothing Then
							If oEmpleado.codigo <> "3943" Then
								For Each xi In oEmpleado.familiares
									If xi.parentesco.codigo = "02" Then
										xCantHijos = xCantHijos + 1
									End If
								Next
							End If
						End If 
						If xCantHijos = 0 Then
							xCantHijos = "00"
						End If 
						If oEmpleado.boextension.adherente > 0 Then
							xCant_Adher = oEmpleado.boextension.adherente
						Else
							xCant_Adher = "00"
						End If  

						'DATOS DE LA CONSULTA
						xCuit          = Formatcuit(cstr(xrst("CUIT").value))
						xNombre        = xrst("NOMBRE_EMPLEADO").value
						xConyuge       = xrst("CONYUGE").value
						xCod_Situacion = xrst("SITUACION").value
						xCod_Condicion = xrst("CONDICION").value
						xCod_Actividad = xrst("CODIGO_ACTIVIDAD").value
						xCod_Zona      = xrst("CODIGO_ZONA").value
						xPorc_Ap_OS    = xrst("PORCENTAJE_APORTE_OBRASOCIAL").value
						xCod_Mod_Cont  = xrst("CODIGO_MODALIDAD_CONTRATACION").value
						xCod_Obra_Soc  = xrst("CODIGO_OBRA_SOCIAL").value
						xRem_Total     = xrst("REMUNERACION_TOTAL").value
						xRem_1         = xrst("REM_1").value
						xAsig_Fam_Pag  = xrst("ASIGNACION_FAMILIAR_PAGA").value
						xImp_Apo_Vol   = xrst("IMPORTE_APORTE_VOLUNTARIO").value
						xImp_Adic_OS   = cDbl(xrst("IMPORTE_ADICIONAL_OBRA_SOCIAL").value)
						xImp_Ex_SS     = xrst("IMPORTE_APORTE_SS").value
						xImp_Ex_OS     = xrst("IMPORTE_EXC_APORTE_OS").value
						xProv_Local    = xrst("LOCALIDAD").value
						xRem_2         = xrst("REM_2").value
						xRem_3         = xrst("REM_3").value
						xRem_4         = xrst("REM_4").value
						xCod_Siniestro = xrst("CODIGO_SINIESTRO").value
						xCorresp_Redu  = xrst("CORRESPONDE_REDUCCION").value
						xCap_Rec_LRT   = xrst("CAPITAL_RECOMPOSICION").value
						xTipoEmpresa   = xrst("TIPO_EMPRESA").value
						xAp_Adic_OS    = xrst("ADICIONAL_OS").value
						xRegimen       = xrst("REGIMEN").value
						xSituacion_R_1 = xrst("SITUACION_1").value
						xDia_Sit_R_1   = xrst("DIA_REVISTA_1").value
						xSituacion_R_2 = xrst("SITUACION_2").value
						xDia_Sit_R_2   = xrst("DIA_REVISTA_2").value
						xSituacion_R_3 = xrst("SITUACION_3").value
						xDia_Sit_R_3   = xrst("DIA_REVISTA_3").value
						xSueldo        = xrst("SUELDO").value
						xSAC           = xrst("SAC").value
						xHoras_Extras  = xrst("HORAS_EXTRAS").value
						xZona_Desfav   = xrst("ZONA_DESFAVORABLE").value
						xVacaciones    = xrst("VACACIONES").value
						xCant_Dia_Trab = cDbl(xrst("DIAS_TRABAJADOS").value)
						xRem_5         = xrst("REM_5").value
						xTrab_Convenc  = xrst("CONVENCIONADO").value
						xRem_6         = xrst("REMUNERACION_6").value
						xTipo_Oper     = xrst("TIPO_OPERACION").value
						xAdicionales   = xrst("ADICIONAL").value
						xPremios       = Round(CDbl(xrst("PREMIO").value), 2)
						xRem_8         = xrst("REM_8").value
						xRem_7         = xrst("REMUNERACION_7").value
						xCant_Hras_Ex  = xrst("CANT_HORAS_EXTRAS").value
						xRem_No_Rem    = xrst("NO_REMUNERATIVO").value
						xMaternidad    = xrst("MATERNIDAD").value
						xRec_Remunera  = xrst("RECTIFICACION").value
						xRem_9         = xrst("REM_9").value
						xCont_Tar_Dife = xrst("CONTRIBUCION_TAREA_DIF").value
						xHoras_Trab    = xrst("HORAS_TRABAJADAS").value
						xSeg_Colectivo = xrst("SEGURO_VIDA").value
                        xImp_Detraer   = xrst("IMPORTE_DETRAER").value
						xInc_Salarial  = 0.00 ' xrst("INCREMENTO_SALARIAL").value
						xRem_11        = 0.00 ' xrst("REM_11").value
						' *************************************************************************************************** '
						If ProgressControlCancelled(Self.WorkSpace) Then
'							Call MsgBox("Proceso Cancelado", 64, "Información")
							call ProgressControlfinish(Self.WorkSpace)
							Exit Sub
						Else
							Call ProgressControlAvance( Self.WorkSpace, "" )
							Call ProgressControlStatusText( Self.WorkSpace, " Contador " & xContador & "  ¦  Empleado: " & xLegajo & " - " & xNombre)
						End If
						' *************************************************************************************************** '
					If UCase(NombreUsuario()) <> "DESARROLLO2" Then
						xSueldoAux = ROUND(cDbl(xSueldo) + cDbl(xSac) + cDbl(xHoras_Extras) + cDbl(xVacaciones) + cDbl(xAdicionales) + cDbl(xPremios), 2)
						If ABS(cDbl(xRem_2)) <> 0.01 Then
							If cDbl(xRem_2) <> xSueldoAux Then
								If cDbl(xRem_2) < xSueldoAux Then
									If (cDbl(xSac) + cDbl(xAdicionales) + cDbl(xPremios) + cDbl(xSueldo) + cDbl(xHoras_Extras) + cDbl(xVacaciones)) > cDbl(xRem_2) + cDbl(xSac) + cDbl(xAdicionales) + cDbl(xPremios) + cDbl(xSueldo) Then
										If cDbl(xSueldo) > 0 Then
											xSueldo = cDbl(xSueldo) - ( xSueldoAux - cDbl(xRem_2) )
										End If
										If cDbl(xSac)         > 0 Then xSac          = 0 End If
										If cDbl(xAdicionales) > 0 Then xAdicionales  = 0 End If
										If cDbl(xPremios)     > 0 Then xPremios      = 0 End If
									ElseIf (cDbl(xSac) + cDbl(xAdicionales) + cDbl(xPremios) + cDbl(xSueldo) + cDbl(xHoras_Extras) + cDbl(xVacaciones)) > cDbl(xRem_2) + cDbl(xSac) + cDbl(xAdicionales) + cDbl(xPremios) Then
										If cDbl(xSac)         > 0 Then xSac          = 0 End If
										If cDbl(xAdicionales) > 0 Then xAdicionales  = 0 End If
										If cDbl(xPremios)     > 0 Then xPremios      = 0 End If
									ElseIf (cDbl(xSac) + cDbl(xAdicionales) + cDbl(xPremios) + cDbl(xSueldo) + cDbl(xHoras_Extras) + cDbl(xVacaciones)) > cDbl(xRem_2) + cDbl(xSac) + cDbl(xAdicionales) Then
										If cDbl(xSac)         > 0 Then xSac          = 0 End If
										If cDbl(xAdicionales) > 0 Then xAdicionales  = 0 End If
									ElseIf (cDbl(xSac) + cDbl(xAdicionales) + cDbl(xPremios) + cDbl(xSueldo) + cDbl(xHoras_Extras) + cDbl(xVacaciones)) > cDbl(xRem_2) + cDbl(xSac) Then
										If cDbl(xSac)         > 0 Then xSac          = 0 End If
									ElseIf (cDbl(xSac) + cDbl(xAdicionales) + cDbl(xPremios) + cDbl(xSueldo) + cDbl(xHoras_Extras) + cDbl(xVacaciones)) > cDbl(xRem_2) + cDbl(xSac) + cDbl(xAdicionales) + cDbl(xPremios) + cDbl(xHoras_Extras) Then
										If cDbl(xSac)         > 0 Then xSac          = 0 End If
										If cDbl(xAdicionales) > 0 Then xAdicionales  = 0 End If
										If cDbl(xPremios)     > 0 Then xPremios      = 0 End If
										If cDbl(xHoras_Extras) > 0 Then
											xHoras_Extras = cDbl(xHoras_Extras) - ( xSueldoAux - cDbl(xRem_2) ) 
										End If
									End If

								Else

								End If
							Else ' SI xRem_2 = xSueldoAux
								If cDbl(xHoras_Extras) > 0 Or cDbl(xVacaciones) > 0 Or cDbl(xSAC) > 0 Then
									If cDbl(xSueldo) < 0 Then ' SI HAY AJUSTE
										If cDbl(xSAC) > 0 Then
											xSueldo  = 0
											If cDbl(xVacaciones) > 0 And cDbl(xHoras_Extras) > 0 And ( xSueldoAux - cDbl(xVacaciones) - cDbl(xHoras_Extras) > 0 ) Then
												xSAC          = xSueldoAux - cDbl(xVacaciones) - cDbl(xHoras_Extras)
											ElseIf cDbl(xVacaciones) > 0 And ( xSueldoAux - cDbl(xVacaciones) > 0 ) Then
												xSAC          = xSueldoAux - cDbl(xVacaciones)
												xHoras_Extras = 0
											ElseIf cDbl(xHoras_Extras) > 0 And ( xSueldoAux - cDbl(xHoras_Extras) > 0 ) Then
												xSAC          = xSueldoAux - cDbl(xHoras_Extras)
												xVacaciones   = 0
											Else
												xSAC          = xSueldoAux
												xHoras_Extras = 0
												xVacaciones   = 0
											End If
											xAdicionales  = 0
											xPremios      = 0
										ElseIf cDbl(xHoras_Extras) > 0 Then
											xSueldo       = 0
											xSac          = 0
											If cDbl(xVacaciones) > 0 And ( xSueldoAux - cDbl(xVacaciones) > 0 ) Then
												xHoras_Extras = xSueldoAux - cDbl(xVacaciones) - 0.01
												xVacaciones   = 0.01
											Else
												xHoras_Extras = xSueldoAux
												xVacaciones   = 0
											End If
											xAdicionales  = 0
											xPremios      = 0
										ElseIf cDbl(xVacaciones) > 0 Then
											xSueldo       = 0
											xSac          = 0
											If cDbl(xHoras_Extras) > 0 And ( xSueldoAux - cDbl(xHoras_Extras) > 0 ) Then
												xHoras_Extras = 0.01
												xVacaciones   = xSueldoAux - cDbl(xHoras_Extras) - 0.01
											Else
												xHoras_Extras = 0
												xVacaciones   = xSueldoAux
											End If
											xAdicionales  = 0
											xPremios      = 0
										End If

										If cDbl(xHoras_Extras) > 0 And cDbl(xVacaciones) > 0 Then
											xSueldo       = 0
											xSac          = 0
											xHoras_Extras = 0.01
											xVacaciones   = cDbl(xRem_2) - 0.01
											xAdicionales  = 0
											xPremios      = 0
										ElseIf cDbl(xHoras_Extras) > cDbl(xRem_2) And cDbl(xVacaciones) = 0 Then ' And  Then
											xSueldo       = 0
											xSac          = 0
											xHoras_Extras = cDbl(xRem_2)
											xVacaciones   = 0
											xAdicionales  = 0
											xPremios      = 0
										ElseIf cDbl(xVacaciones) > cDbl(xRem_2) And cDbl(xHoras_Extras) = 0 Then
											xSueldo       = 0
											xSac          = 0
											xHoras_Extras = 0
											xVacaciones   = cDbl(xRem_2)
											xAdicionales  = 0
											xPremios      = 0
										ElseIf cDbl(xVacaciones) > 0 And cDbl(xHoras_Extras) = 0 Then
											
										End If
									End If
								Else
									If cDbl(xSueldo) < 0 Then
										xSueldo      = cDbl(xRem_2)
										xAdicionales = 0
										xPremios     = 0
									End If
								End If
							End If
						Else

							xSueldo       = 0.01
							xSac          = 0
							xHoras_Extras = 0
							xVacaciones   = 0
							xAdicionales  = 0
							xPremios      = 0
						End If
					Else
						' *************************************************************************************************** '
                        xSueldo2 = cDbl(xSueldo) + cDbl(xHoras_Extras) + cDbl(xVacaciones) + cDbl(xAdicionales) + cDbl(xPremios)
                        xSueldo  = xSueldo2
'						Comentado por MLeon (20200121)
'						If ( Round(cDbl(xRem_2), 2) = Round(cDbl(xSueldo2) + cDbl(xSac), 2) And cDbl(xHoras_Extras) = 0 ) Or _
'						   cDbl(xRem_2) = 0.01 Then
'							xHoras_Extras = 0
'							xVacaciones   = 0
'							xAdicionales  = 0
'							xPremios      = 0
'						End If

						'CONTROL DE APLICACION DE AJUSTE EN SUELDO O SAC PROPORCIONAL
						'Sac Proporcional (Liq Final):  Se da una situacion con los ajustes negativos. 
						'Si el ajuste excede el sueldo, la diferencia se debe descontar del sac (6773 - godoy)
						If cDbl(xSueldo) < 0 And cDbl(xSac) > 0 Then
							xSac    = cDbl(xSac) + cDbl(xSueldo)
							xSueldo = 0
						End If

						'CONTROL PARA SUELDO + ADICIONALES Y REMUNERACION 2
						'ESTO ES PORQUE EL ADICIONAL Y EL SUELDO DAN CON UN CENTAVO DE DIFERENCIA - CUANDO EL EMPLEADO TIENE 30 DIAS DESCONTADO
						If abs(cDbl(xSueldo)) = abs((cDbl(xAdicionales)-0.01)) Then 
							xSueldoAdicional = 0
						Else
							xSueldoAdicional = cDbl(xSueldo) + cDbl(xAdicionales) + cDbl(xPremios) + cDbl(xSac)
							If xSueldoAdicional <= 0.01 Then ' xSueldoAdicional < 0
								xSueldoAdicional	= 0 
								xSueldo				= 0 
								xAdicionales		= 0 
								xPremios			= 0 
								xSac				= 0 
'								xVacaciones			= 0 
							End If 
						End If
						If xSueldoAdicional = 0 and cDbl(xVacaciones) > 0 Then
							xSueldo			= 0 
							xAdicionales	= 0 
							xPremios		= 0
							xVacaciones		= cDbl(xVacaciones) - 0.01 
						End If 
						If xSueldoAdicional = 0 and cDbl(xSAC) > 0 Then
							xSueldo			= 0 
							xAdicionales	= 0 
							xPremios		= 0
							xSAC			= cDbl(xSAC) - 0.01 
						End If
						If xSueldoAdicional = 0 Then 
							xSueldo			= 0 
							xAdicionales	= 0 
							xPremios		= 0
						End If
					End If
						' *************************************************************************************************** '

						'CONTROL DE CANTIDAD DE HORAS EXTRAs
						If cDbl(xCant_Hras_Ex) > 0 and cDbl(xCant_Hras_Ex) < 1 Then
							xCant_Hras_Ex = 1
						End If 

						'LLENO LOS CAMPOS
						xcampopro(1) = string(11-len(cstr(xCuit))," ")&cstr(xCuit)
						xcampopro(2) = cstr(left(cstr(xNombre)&"                              ", 30))
						xcampopro(3) = string(1-len(cstr(xConyuge)),"0")&cstr(xConyuge)
						xcampopro(4) = string(2-len(cstr(xCantHijos)),"0")&cstr(xCantHijos)
						xcampopro(5) = string(2-len(cstr(xCod_Situacion)),"0")&cstr(xCod_Situacion)
						xcampopro(6) = string(2-len(cstr(xCod_Condicion)),"0")&cstr(xCod_Condicion)
						xcampopro(7) = string(3-len(cstr(xCod_Actividad)),"0")&cstr(xCod_Actividad)
						xcampopro(8) = string(2-len(cstr(xCod_Zona)),"0")&cstr(xCod_Zona)
						xcampopro(9) = string(5-len(cstr(xPorc_Ap_OS)),"0")&cstr(xPorc_Ap_OS)
						xcampopro(10)= string(3-len(cstr(xCod_Mod_Cont)),"0")&cstr(xCod_Mod_Cont)
						xcampopro(11)= string(6-len(cstr(xCod_Obra_Soc)),"0")&cstr(xCod_Obra_Soc)
						xcampopro(12)= string(2-len(cstr(xCant_Adher)),"0")&cstr(xCant_Adher)
						xcampopro(13)= string(12-(len(Formatcuit1(Formatnumber(xRem_Total,2)))),"0")&Formatcuit1(Formatnumber(xRem_Total,2))
						xcampopro(14)= string(12-(len(Formatcuit1(Formatnumber(xRem_1,2)))),"0")&Formatcuit1(Formatnumber(xRem_1,2))
						xcampopro(15)= string(9-len(cstr(xAsig_Fam_Pag)),"0")&cstr(xAsig_Fam_Pag)
						xcampopro(16)= string(9-len(cstr(xImp_Apo_Vol)),"0")&cstr(xImp_Apo_Vol)
						If CDbl(xImp_Adic_OS) = 0 Then
							xcampopro(17) = "000000,00"
						Else
							xcampopro(17) = string(9-(len(Formatcuit1(Formatnumber(xImp_Adic_OS,2)))),"0")&Formatcuit1(Formatnumber(xImp_Adic_OS,2))
						End If
'						xcampopro(18)= string(9-len(cstr(xImp_Ex_SS)),"0")&cstr(xImp_Ex_SS)
						xcampopro(18)= string(9-(len(Formatcuit1(Formatnumber(xImp_Ex_SS,2)))),"0")&Formatcuit1(Formatnumber(xImp_Ex_SS,2))
						
						xcampopro(19)= string(9-len(cstr(xImp_Ex_OS)),"0")&cstr(xImp_Ex_OS)
						xcampopro(20)= cstr(left(cstr(xProv_Local)&"                                                  ", 50))
						xcampopro(21)= string(12-(len(Formatcuit1(Formatnumber(xRem_2,2)))),"0")&Formatcuit1(Formatnumber(xRem_2,2))
						xcampopro(22)= string(12-(len(Formatcuit1(Formatnumber(xRem_3,2)))),"0")&Formatcuit1(Formatnumber(xRem_3,2))
						xcampopro(23)= string(12-(len(Formatcuit1(Formatnumber(xRem_4,2)))),"0")&Formatcuit1(Formatnumber(xRem_4,2))
						xcampopro(24)= string(2-len(cstr(xCod_Siniestro)),"0")&cstr(xCod_Siniestro)
						xcampopro(25)= string(1-len(cstr(xCorresp_Redu)),"0")&cstr(xCorresp_Redu)
						xcampopro(26)= string(9-len(cstr(xCap_Rec_LRT)),"0")&cstr(xCap_Rec_LRT)
						xcampopro(27)= string(1-len(cstr(xTipoEmpresa)),"0")&cstr(xTipoEmpresa)
						If CDbl(xAp_Adic_OS) = 0 Then
							xcampopro(28)= "000000,00"
						Else
							xcampopro(28)= string(9-(len(Formatcuit1(Formatnumber(xAp_Adic_OS,2)))),"0")&Formatcuit1(Formatnumber(xAp_Adic_OS,2)) 
						End If
						xcampopro(29)= string(1-len(cstr(xRegimen)),"0")&cstr(xRegimen)
						xcampopro(30)= string(2-len(cstr(xSituacion_R_1)),"0")&cstr(xSituacion_R_1)
						xcampopro(31)= string(2-len(cstr(xDia_Sit_R_1)),"0")&cstr(xDia_Sit_R_1)
						xcampopro(32)= string(2-len(cstr(xSituacion_R_2)),"0")&cstr(xSituacion_R_2)
						xcampopro(33)= string(2-len(cstr(xDia_Sit_R_2)),"0")&cstr(xDia_Sit_R_2)
						xcampopro(34)= string(2-len(cstr(xSituacion_R_3)),"0")&cstr(xSituacion_R_3)
						xcampopro(35)= string(2-len(cstr(xDia_Sit_R_3)),"0")&cstr(xDia_Sit_R_3)
					   
						If cDbl(xSueldo) = 0 And cDbl(xRem_2) = 0 Then
							xcampopro(36)= "000000000,01"
						Else
							xSueldo = cDbl(xRem_2) - (cDbl(xSAC) + cDbl(xHoras_Extras) + cDbl(xZona_Desfav) + cDbl(xVacaciones) + cDbl(xAdicionales) + cDbl(xPremios))
							xcampopro(36)= string(12-(len(Formatcuit1(Formatnumber(xSueldo,2)))),"0")&Formatcuit1(Formatnumber(xSueldo,2))
						End If

						If cDbl(xSAC) = 0 Then
							xcampopro(37)= "000000000,00"
						Else
							xcampopro(37)= string(12-(len(Formatcuit1(Formatnumber(xSAC,2)))),"0")&Formatcuit1(Formatnumber(xSAC,2))
						End If

						If cDbl(xHoras_Extras) = 0 Then
							xcampopro(38)= "000000000,00"
						Else
							xcampopro(38)= string(12-(len(Formatcuit1(Formatnumber(xHoras_Extras,2)))),"0")&Formatcuit1(Formatnumber(xHoras_Extras,2))
						End If

						xcampopro(39)= string(12-len(cstr(xZona_Desfav)),"0")&cstr(xZona_Desfav)

						If cDbl(xVacaciones) = 0 Then
							xcampopro(40)= "000000000,00"
						Else
							xcampopro(40)= string(12-(len(Formatcuit1(Formatnumber(xVacaciones,2)))),"0")&Formatcuit1(Formatnumber(xVacaciones,2))
						End If
					   
						If xCant_Dia_Trab = 0.01 Then
							xcampopro(41)="000001,00"
						ElseIf xCant_Dia_Trab >= 30 Then
							xcampopro(41)="000030,00"
						Else
							xcampopro(41)=STRING(9-(LEN(Formatcuit1(Formatnumber(xCant_Dia_Trab,2)))),"0")&Formatcuit1(Formatnumber(xCant_Dia_Trab,2)) 
						End If
						xcampopro(42)= string(12-(len(Formatcuit1(Formatnumber(xRem_5,2)))),"0")&Formatcuit1(Formatnumber(xRem_5,2))
						xcampopro(43)= string(1-len(cstr(xTrab_Convenc)),"0")&cstr(xTrab_Convenc)
						xcampopro(44)= string(12-len(cstr(xRem_6)),"0")&cstr(xRem_6)
						xcampopro(45)= string(1-len(cstr(xTipo_Oper)),"0")&cstr(xTipo_Oper)

						If cDbl(xAdicionales) = 0 Then
							xcampopro(46)= "000000000,00"
						Else
							xcampopro(46)= string(12-(len(Formatcuit1(Formatnumber(xAdicionales,2)))),"0")&Formatcuit1(Formatnumber(xAdicionales,2))
						End If

						If cDbl(xPremios) = 0 Then
							xcampopro(47)= "000000000,00"
						Else
							xcampopro(47)= string(12-(len(Formatcuit1(Formatnumber(xPremios,2)))),"0")&Formatcuit1(Formatnumber(xPremios,2))
						End If 

						xcampopro(48)= string(12-(len(Formatcuit1(Formatnumber(xRem_8,2)))),"0")&Formatcuit1(Formatnumber(xRem_8,2))
						xcampopro(49)= string(12-len(cstr(xRem_7)),"0")&cstr(xRem_7)
						xcampopro(50)= string(3-len(cint(xCant_Hras_Ex)),"0")&cint(xCant_Hras_Ex)
					   
						If cDbl(xRem_No_Rem) = 0.01 or cDbl(xRem_No_Rem) = 0 Then
							xcampopro(51)= "000000000,00"
						Else
							xcampopro(51)= string(12-(len(Formatcuit1(Formatnumber(xRem_No_Rem,2)))),"0")&Formatcuit1(Formatnumber(xRem_No_Rem,2))
						End If 

						xcampopro(52)= string(12-(len(Formatcuit1(Formatnumber(xMaternidad,2)))),"0")&Formatcuit1(Formatnumber(xMaternidad,2))
						If xcampopro(52) = "000000000,01" Then
							xcampopro(52) = "000000000,00"
						End If
						xcampopro(53)= string(9-len(cstr(xRec_Remunera)),"0")&cstr(xRec_Remunera)
						xcampopro(54)= string(12-(len(Formatcuit1(Formatnumber(xRem_9,2)))),"0")&Formatcuit1(Formatnumber(xRem_9,2))
						xcampopro(55)= string(9-len(cstr(xCont_Tar_Dife)),"0")&cstr(xCont_Tar_Dife)
                        If xHoras_Trab = 0.00 Then
'                            If xCant_Dia_Trab = 0.01 Then
'                                xcampopro(56) = "008"
'                            Else
'                                xcampopro(56) = "000"
'                            End If						
							If oEmpleado.Perfil.Descripcion = "U.O.M." Then
							    xcampopro(56) = "008"
                            Else
                                xcampopro(56) = "000"
							End If
						Else
							xcampopro(56)= string(3-len(cstr(xHoras_Trab)),"0")&cstr(xHoras_Trab)
						End If
						xcampopro(57)= string(1-len(cstr(xSeg_Colectivo)),"0")&cstr(xSeg_Colectivo)
                        xcampopro(58)= string(12-(len(Formatcuit1(Formatnumber(xImp_Detraer,2)))),"0")&Formatcuit1(Formatnumber(xImp_Detraer,2))
						xcampopro(59)= string(12-(len(Formatcuit1(Formatnumber(xIncremento_Solidario,2)))),"0")&Formatcuit1(Formatnumber(xIncremento_Solidario,2))
						xcampopro(60)= string(12-(len(Formatcuit1(Formatnumber(xRem_11,2)))),"0")&Formatcuit1(Formatnumber(xRem_11,2))
						xstring = ""
						For I = 1 to 60  ' 58
							xstring = xstring & xcampopro(I)
							If xExcel Then
                                If I = 1 Then
                                    xstring2 = xcampopro(I)
                                Else
								    xstring2 = xstring2 & ";" & xcampopro(I)
                                End If
                            End If
			       		Next
                        If xExcel Then
							HojaExcel.ActiveSheet.Cells(fila, 1).Value = xstring2
						End If
						Fila = Fila + 1
						xstring = xstring
		  	 	   		Call EscribirTXT( Path & xNombreTxt, xstring , 1 )
		  	 	   		xrst.moveNext
						xContador = xContador + 1
		    wend
			If xExcel Then
				' Autofit
				HojaExcel.ActiveSheet.Columns("A:Z").AutoFit
				' Guardo excel
				HojaExcel.ActiveWorkbook.SaveAs Path & xNombreXls
				HojaExcel.Visible = True
			End If
			Call ProgressControlStatusText( Self.WorkSpace, "Empleados Procesados: " & xContador - 1 & vbCrlf & "  -  Proceso Finalizado.")
			Call MsgBox(" EL PROCESO FINALIZÓ CORRECTAMENTE", 64, "Información")
			Call ProgressControlFinish( Self.WorkSpace )
    Else
 	     msgbox"El Proceso ha Sido Cancelado por el Usuario"
	     Exit Sub
    End If
End Sub

function Formatcuit(cCuit)
	lencuit=len(cCuit)
   	For j = 1 to lencuit 
       cNumero = mid(cCuit,j,1)
       If IsNumeric(cNumero) Then 
          cCUITLimpio = cCUITLimpio + cNumero
	   ElseIf cNumero<>"." and  cNumero="," Then
	      cCUITLimpio = cCUITLimpio + cNumero
       End If
    Next 
	Formatcuit = cCUITLimpio
end function

function Formatcuit1(cCuit)
    cCuit = Formatnumber(cdbl(cCuit),2)
 	If "1.1"=CDBL("1.1") Then
	      cCuit=Replace(cCuit, ",", "")
	      cCuit=Replace(cCuit, ".", ",")
	Else 
	      cCuit=Replace(cCuit, ".", "")
	End If
	lencuit=len(cCuit)
   	For j = 1 to lencuit 
       cNumero = mid(cCuit,j,1)
       If IsNumeric(cNumero) Then 
          cCUITLimpio = cCUITLimpio + cNumero
	   ElseIf cNumero<>"." and  cNumero="," Then
	      cCUITLimpio = cCUITLimpio + cNumero
       End If
    Next 
	Formatcuit1 = cCUITLimpio
end function

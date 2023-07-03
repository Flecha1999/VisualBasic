REM ACUTALIZAR CATEGORIA
sub main()
	Stop
	Set xView1 = NewCompoundView(self,"Perfil", self.Workspace, nil, True)
	Set xColumnum  = xView1.ColumnFromPath( "descripcion" )
	
	Set xVisualVar = VisualVarEditor ("Actualizar Convenio")
	Call AddVarView   (xVisualVar, "Convenio", "Seleccione Convenio","Convenio" , xView1, "descripcion")	
    'ultimoDiaDelMes= Date.DaysInMonth(year(date()), Month(date()))
	'Dim daysInJuly As Integer = System.DateTime.DaysInMonth(year(date), month(date))	
    'dateserial(year(now()),month(now)-1,1 )
	Call AddVarDate( xvisualvar, "Fecha", "Fecha", "Fecha",dateserial(year(now()),month(now)-1,1 ) )

	xAceptar = ShowVisualVar (xVisualVar)
	If xAceptar Then
		convenio = GetValueVisualVar (xVisualVar,"Convenio", "Convenio")
		xdesdefecha = cdate(GetValueVisualVar( xvisualvar, "Fecha", "Fecha" ))
		strdesdefecha = year(xdesdefecha) & right("00" & month(xdesdefecha),2) & right("00" & day(xdesdefecha),2) 
	Else
		msgbox"Proceso Cancelado por el usuario"
	End If
	
    xstring=stringconexion("Calipso",self.workspace) : Set xcone = createobject("adodb.connection")
	xcone.connectiontimeout = 4500 : xcone.connectionstring = xstring : xcone.open	 
	Contador = 0
	If convenio <> Empty Then
		If convenio = "{CF65CADD-83AA-43BF-AED4-CF7071A2A061}" Then 'som
			sql =	" SELECT  ID, categoria_id " &_
					"   FROM EMPLEADO " &_
					" WHERE PERFIL_ID = 'CF65CADD-83AA-43BF-AED4-CF7071A2A061' " &_
					"     AND (DATEDIFF(" & "M" & ",FECHAINGRESO,'" & strdesdefecha & "'))>=15 " &_
					"     AND CATEGORIA_ID IN ('A6EB2A10-8445-4D8C-9AA2-5DBC2C6299D4', 'A10E767B-D318-487C-9139-3B7B81E820C5', 'C0CE123C-D812-4858-A033-2884112BD765') " &_
					"     AND ACTIVESTATUS = '0' "
			
			Set xrst = RecordSet(xCone, "select top 1 * from producto")
			xrst.close
			xrst.activeconnection.commandtimeout=0
			xrst.source=sql
			xrst.open
			While xrst.eof = False 
				Set xempleado = ExisteBo(self,"empleado","id",xrst("id").value, nil,true,False,"=")
				Select Case xrst("categoria_id").value
					Case "{A6EB2A10-8445-4D8C-9AA2-5DBC2C6299D4}"		' SOM - Oficial
						Set xcategoria = ExisteBo(Self, "CATEGORIA", "ID", "94A5EA46-BD06-4C09-BF7B-A26AA17616AE", nil, True, False,"=")	' SOM - Oficial 1era
					Case "{C0CE123C-D812-4858-A033-2884112BD765}"		' SOM - Oficial Jornada Reducida Iscot
						Set xcategoria = ExisteBo(Self, "CATEGORIA", "ID", "219BB85B-2EED-4B63-9298-F9800D1F2C3A", nil, True, False,"=")	' SOM - Oficial 1era Jornada Reducida Iscot
					Case "{A10E767B-D318-487C-9139-3B7B81E820C5}"		' SOM - Oficial Jornada Reducida Convenio
						Set xcategoria = ExisteBo(Self, "CATEGORIA", "ID", "13B48247-0DD8-403F-A7D2-DAA17CA5E289", nil, True, False,"=")	' SOM - Oficial 1era Jornada Reducida Convenio
				End Select
				If Not xcategoria Is Nothing Then
				   xempleado.categoria = xcategoria
				   Call workspacecheck(self.workspace)
				   Call sendDebug( "Actualizó   --  " & xempleado.descripcion )
				   Contador = Contador + 1
				Else
				   Call sendDebug( "No se pudo Actualizar   --  " & xempleado.descripcion )
				End If
				xrst.movenext
		    Wend 

		ElseIf convenio = "{BDD8FB00-8B97-4C7E-A45A-C637543598B3}" Then 'UOM
            ' A los 6 meses de antiguedad pasan de ingresante a operario calificado ---- de acuerdo a la rama
			SQL =	" SELECT ID, CATEGORIA_ID, DESCRIPCION, FECHAINGRESO, " &_
					" 	     CAST(YEAR(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR) + RIGHT('00' + CAST(MONTH(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR), 2) + RIGHT('00' + CAST(DAY(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR), 2) AS FINCAMBIO " &_
					"   FROM EMPLEADO " &_
					"  WHERE PERFIL_ID = '{BDD8FB00-8B97-4C7E-A45A-C637543598B3}' " &_
					"    AND CAST(YEAR(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR) + RIGHT('00' + CAST(MONTH(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR), 2) + RIGHT('00' + CAST(DAY(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR), 2) <= '" & strdesdefecha & "' " &_
					"    AND CATEGORIA_ID IN ('BB7A80A0-6883-4D7A-AB19-153AB0CE3A6B', 'CD5A1CBE-A4CC-4C78-96D6-DB7BE8B9D05D') " &_
					"    AND ACTIVESTATUS = '0' " &_
					"  ORDER BY 3 "

			Set xrst = RecordSet(xCone, "select top 1 * from producto")
			xrst.close
			xrst.activeconnection.commandtimeout=0
			xrst.source=sql
			xrst.open	
			While xrst.eof = False 
				Set xempleado = ExisteBo(self,"empleado","id",xrst("id").value, nil,true,False,"=")
				Select Case xrst("categoria_id").value
					Case "{CD5A1CBE-A4CC-4C78-96D6-DB7BE8B9D05D}"		' UOM - Rama 4 - Laudo 29 - Ingresante	UOM R4 - L29 - Ingre
						Set xcategoria = ExisteBo(Self, "CATEGORIA", "ID", "5D6092BF-1BA4-4B9C-8DFF-6BCC5C0D4DC1", nil, True, False,"=")	' UOM - Rama 4 - Laudo 29 - Operario Calificado	UOM R4 - L29 - Op Ca
					Case "{BB7A80A0-6883-4D7A-AB19-153AB0CE3A6B}"		' UOM - Rama 17 - Ingresante	UOM R17 - Ingresante
						Set xcategoria = ExisteBo(Self, "CATEGORIA", "ID", "AB10EFE2-5677-497A-8D9E-F5BE29DE41A4", nil, True, False,"=")	' UOM - Rama 17 - Operario Calificado	UOM R17 - Operario C
				End Select
				If Not xcategoria Is Nothing Then
				   xempleado.categoria = xcategoria
				   Call workspacecheck(self.workspace)
				   Call sendDebug( "Actualizó   --  " & xempleado.descripcion )
				   Contador = Contador + 1
				Else
				   Call sendDebug( "No se pudo Actualizar   --  " & xempleado.descripcion )
				End If
				xrst.movenext
		    Wend 
		ElseIf convenio = "{6DC95932-42F4-40A1-ABE3-69A680AA7A5A}" Then 'SMATA MENSUAL
            '------------la categoria smata vw inicial y cuando pasan 6 meses pasar a smata vw categoria 1 ---- de acuerdo a la rama
			SQL =	" SELECT ID, CATEGORIA_ID, DESCRIPCION, FECHAINGRESO, " &_
					" 	     CAST(YEAR(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR) + RIGHT('00' + CAST(MONTH(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR), 2) + RIGHT('00' + CAST(DAY(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR), 2) AS FINCAMBIO " &_
					"   FROM EMPLEADO " &_
					"  WHERE PERFIL_ID = '{6DC95932-42F4-40A1-ABE3-69A680AA7A5A}' " &_
					"    AND CAST(YEAR(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR) + RIGHT('00' + CAST(MONTH(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR), 2) + RIGHT('00' + CAST(DAY(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR), 2) <= '" & strdesdefecha & "' " &_
					"    AND CATEGORIA_ID IN ('09E10863-4C12-48AB-A5BD-861B4E6A008A') " &_
					"    AND ACTIVESTATUS = '0' " &_
					"  ORDER BY 3 "

			Set xrst = RecordSet(xCone, "select top 1 * from producto")
			xrst.close
			xrst.activeconnection.commandtimeout=0
			xrst.source=sql
			xrst.open	
			While xrst.eof = False 
				Set xempleado = ExisteBo(self,"empleado","id",xrst("id").value, nil,true,False,"=")
				Select Case xrst("categoria_id").value
					Case "{09E10863-4C12-48AB-A5BD-861B4E6A008A}"' SMATA VW INICIAL
						Set xcategoria = ExisteBo(Self, "CATEGORIA", "ID", "4F990F0F-71D9-4B97-82E3-AD15E83C5B31", nil, True, False,"=") 'SMATA VW CATEGORIA 1
						End Select
				If Not xcategoria Is Nothing Then
				   xempleado.categoria = xcategoria
				   Call workspacecheck(self.workspace)
				   Call sendDebug( "Actualizó   --  " & xempleado.descripcion )
				   Contador = Contador + 1
				Else
				   Call sendDebug( "No se pudo Actualizar   --  " & xempleado.descripcion )
				End If
				xrst.movenext
		    Wend 

	    ElseIf convenio = "{8A8D8A26-03BF-422D-86E4-69F6E0E0DF62}" Then 'sorbyl
			sql = "select  id from empleado where perfil_id='8A8D8A26-03BF-422D-86E4-69F6E0E0DF62'  and (datediff("&"M"&",fechaingreso,'"&strdesdefecha&"'))>=3 and categoria_id='D0C3598E-153B-47E5-8FE1-A1A695A74D45' and activestatus='0' " 'inicial limpieza
			Set xrst = RecordSet(xCone, "select top 1 * from producto")
			xrst.close
			xrst.activeconnection.commandtimeout=0
			xrst.source=sql
			xrst.open

			While xrst.eof = False 
				Set xempleado=existebo(self,"empleado","id",xrst("id").value, nil,true,False,"=")
				Set xcategoria=existebo(self,"categoria","id","99D45D8B-EF5C-49AF-AF05-7EC545585E9D", nil,true,False,"=") ' limpieza gral
				xempleado.categoria=xcategoria
				Call workspacecheck(self.workspace)
				Call sendDebug( "Actualizó" & xempleado.descripcion )
				Contador = Contador + 1
				xrst.movenext
			Wend 
            '------------------------ SMATA GM MENSUAL
        ElseIf convenio = "{6DC95932-42F4-40A1-ABE3-69A680AA7A5A}" Then 'smata
			' A los 3 meses de antiguedad pasan de smata gm inicial a smata gm operario Generico ---- de acuerdo a la rama
			SQL =	" SELECT ID, CATEGORIA_ID, DESCRIPCION, FECHAINGRESO, " &_
					" CAST(YEAR(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR) + RIGHT('00' + CAST(MONTH(DATEADD(m, 3, FECHAINGRESO)) AS VARCHAR), 2) + RIGHT('00' + CAST(DAY(DATEADD(m, 3, FECHAINGRESO)) AS VARCHAR), 2) AS FINCAMBIO " &_
					"  FROM EMPLEADO " &_
					"  WHERE PERFIL_ID = '{6DC95932-42F4-40A1-ABE3-69A680AA7A5A}' " &_
					"  AND CAST(YEAR(DATEADD(m, 6, FECHAINGRESO)) AS VARCHAR) + RIGHT('00' + CAST(MONTH(DATEADD(m, 3, FECHAINGRESO)) AS VARCHAR), 2) + RIGHT('00' + CAST(DAY(DATEADD(m, 3, FECHAINGRESO)) AS VARCHAR), 2) <= '" & strdesdefecha & "' " &_
					"  AND CATEGORIA_ID IN ('781A56EF-B9D2-4B38-8007-B75EA6E5CC54') " &_
					"  AND ACTIVESTATUS = '0' " &_
					"  ORDER BY 3 "

			Set xrst = RecordSet(xCone, "select top 1 * from producto")
			xrst.close
			xrst.activeconnection.commandtimeout=0
			xrst.source=sql
			xrst.open	
			While xrst.eof = False 
				Set xempleado = ExisteBo(self,"empleado","id",xrst("id").value, nil,true,False,"=")
				Select Case xrst("categoria_id").value
					Case "{781A56EF-B9D2-4B38-8007-B75EA6E5CC54}"		' SMATA GM OPERARIO GENERICO
						Set xcategoria = ExisteBo(Self, "CATEGORIA", "ID", "80B9B4BF-1321-4B9A-8A04-2385BD175C59", nil, True, False,"=")
				End Select
				If Not xcategoria Is Nothing Then
				   xempleado.categoria = xcategoria
				   Call workspacecheck(self.workspace)
				   Call sendDebug( "Actualizó   --  " & xempleado.descripcion )
				   Contador = Contador + 1
				Else
				   Call sendDebug( "No se pudo Actualizar   --  " & xempleado.descripcion )
				End If
				xrst.movenext
		    Wend 
		End If
		MsgBox "Fin de Proceso" & Chr(13) & Chr(13) & " ( " & Contador & " ) Empleados Actualizados"
		xcone.close
	End If
End sub

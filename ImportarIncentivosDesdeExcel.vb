
'CORREGIDO 
Sub Main ()
	Stop
	xPath = OpenFileDialog("C:\", "Archivo (*.xls)")
	If xPath = "" Then
		MsgBox "No seleccionó ningún archivo. El proceso se cancela.", 48, "Aviso" : Exit Sub
	End If
	' -- EXCEL
	Set HojaExcel = CreateObject("Excel.Application") : HojaExcel.Workbooks.Open xPath : HojaExcel.Sheets("Hoja1").Select
	R = 2 : xTotalFilas = 0
	Do While Trim(HojaExcel.ActiveSheet.Cells(R, 1).Value) <> ""
	    xTotalFilas = xTotalFilas + 1 : R = R + 1
	Loop
	Call ProgressControl(Self.Workspace, "Importación de Adicionales desde Excel", 0, xTotalFilas)
	R = 2
	conErrores 	= False
	Do While Trim(HojaExcel.ActiveSheet.Cells(R, 1).Value) <> ""
		Call ProgressControlAvance(Self.Workspace, "Fila: " & R & " - Legajo: " & Trim(HojaExcel.ActiveSheet.Cells(R, 1).Value))
		If HojaExcel.ActiveSheet.Cells(R, 15).Value <> "Empleado Actualizado Correctamente." Then ' Or IsEmpty(HojaExcel.ActiveSheet.Cells(R, 10).Value) Then
			xMsgError = "" : xLegajo = "" : xAdicionalSalud = "" : xAdicionalTareasPel = "" : xAdicioanlTrabajoEqu = ""
			Set oAdicPlaAut = Nothing : Set oIncXPuesto = Nothing : Set oIncXTarea = Nothing : conErrores = False
			xLegajo	= Trim(HojaExcel.ActiveSheet.Cells(R, 1).Value)
			Set xViewEmpleado = NewCompoundView(Self,"EMPLEADO", Self.Workspace, nil, True)
			xViewEmpleado.addfilter( NewFilterSpec(NewColumnSpec( "EMPLEADO", "ACTIVESTATUS", "EMPLEADO" ),"=", 0))
			xViewEmpleado.addfilter( NewFilterSpec(NewColumnSpec( "EMPLEADO", "CODIGO", "EMPLEADO" ),"=", xLegajo))
			For Each ee In xViewEmpleado.ViewItems
				Set oEmpleado = ee.bo : Exit For
			Next
			If Not oEmpleado Is Nothing Then
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalSalud		 = Trim(HojaExcel.ActiveSheet.Cells(R, 2).Value)
				If xAdicionalSalud <> "" Then
					If xAdicionalSalud = "SI" Then oEmpleado.BoExtension.Adicional13 = True Else oEmpleado.BoExtension.Adicional13 = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicioanlTrabajoEqu = Trim(HojaExcel.ActiveSheet.Cells(R, 3).Value)
				If xAdicioanlTrabajoEqu <> "" Then
					If xAdicioanlTrabajoEqu = "SI" Then oEmpleado.BoExtension.Adicional28 = True Else oEmpleado.BoExtension.Adicional28 = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalTareasPel	 = Trim(HojaExcel.ActiveSheet.Cells(R, 4).Value)
				If xAdicionalTareasPel <> "" Then
					If xAdicionalTareasPel = "SI" Then oEmpleado.BoExtension.AdicionalTareasPeligrosas = True Else oEmpleado.BoExtension.AdicionalTareasPeligrosas = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalPlantaAuto = Trim(HojaExcel.ActiveSheet.Cells(R, 5).Value)
				If xAdicionalPlantaAuto <> "" And xAdicionalPlantaAuto <> "BORRAR" Then
					If xAdicionalPlantaAuto = "BORRAR" Then
						oEmpleado.BoExtension.AdicionalPlantaAutomotriz = Nothing 
					Else
						Set xViewAPA = NewCompoundView(Self, "UD_INCENTIVOS", Self.Workspace, nil, True)
						xViewAPA.addfilter( NewFilterSpec(xViewAPA.ColumnFromPath("NOMBRE"),"=", xAdicionalPlantaAuto)) 
						xViewAPA.addfilter( NewFilterSpec(xViewAPA.ColumnFromPath("CONVENIO"),"=", oEmpleado.Perfil))  
						For Each oIt In xViewAPA.ViewItems
							Set oAdicPlaAut = oIt.bo : Exit For
						Next
						If Not oAdicPlaAut Is Nothing Then
							oEmpleado.BoExtension.AdicionalPlantaAutomotriz = oAdicPlaAut 
						Else
							xMsgError = xMsgError & " - " & "No se encontró el Incentivo de Planta Automotriz" : conErrores = True
						End If
					End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xIncentivoXPuesto	 = Trim(HojaExcel.ActiveSheet.Cells(R, 6).Value)
				If xIncentivoXPuesto <> "" And xIncentivoXPuesto <> "BORRAR" Then
					If xIncentivoXPuesto = "BORRAR" Then
						oEmpleado.BoExtension.IncentivoPorPuesto = Nothing 
					Else
						Set xViewAPA = NewCompoundView(Self, "UD_INCENTIVOS", Self.Workspace, nil, True)
						xViewAPA.addfilter( NewFilterSpec(xViewAPA.ColumnFromPath("NOMBRE"),"=", xIncentivoXPuesto)) 
						xViewAPA.addfilter( NewFilterSpec(xViewAPA.ColumnFromPath("CONVENIO"),"=", oEmpleado.Perfil))
						For Each oIt In xViewAPA.ViewItems
							Set oIncPuesto = oIt.bo : Exit For
						Next
						If Not oIncPuesto Is Nothing Then
							oEmpleado.BoExtension.IncentivoPorPuesto = oIncPuesto
						Else
							xMsgError = xMsgError & " - " & "No se encontró el Incentivo por Puesto" : conErrores = True
						End If
					End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xIncentivoXTarea	 = Trim(HojaExcel.ActiveSheet.Cells(R, 7).Value) 
				If xIncentivoXTarea <> "" And xIncentivoXTarea <> "BORRAR" Then
					If xIncentivoXTarea = "BORRAR" Then
						oEmpleado.BoExtension.IncentivoPorTarea = Nothing 
					Else
						Set xViewAPA = NewCompoundView(Self, "UD_INCENTIVOS", Self.Workspace, nil, True)
						xViewAPA.addfilter( NewFilterSpec(xViewAPA.ColumnFromPath("NOMBRE"),"=", xIncentivoXTarea)) 
						xViewAPA.addfilter( NewFilterSpec(xViewAPA.ColumnFromPath("CONVENIO"),"=", oEmpleado.Perfil)) 
						For Each oIt In xViewAPA.ViewItems
							Set oIncTarea = oIt.bo : Exit For
						Next
						If Not oIncTarea Is Nothing Then
							oEmpleado.BoExtension.IncentivoPorTarea = oIncTarea 
						Else
							xMsgError = xMsgError & " - " & "No se encontró el Incentivo por Tarea" : conErrores = True
						End If
					End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalEspProdPlantaAutomotriz = Trim(HojaExcel.ActiveSheet.Cells(R, 8).Value)
				If xAdicionalEspProdPlantaAutomotriz <> "" Then
					If xAdicionalEspProdPlantaAutomotriz = "SI" Then oEmpleado.BoExtension.AdicEspProdPlantaAuto = True Else oEmpleado.BoExtension.AdicEspProdPlantaAuto = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalPlantaAutopartista = Trim(HojaExcel.ActiveSheet.Cells(R, 9).Value)
				If xAdicionalPlantaAutopartista <> "" Then
					If xAdicionalPlantaAutopartista = "SI" Then oEmpleado.BoExtension.ADICIONALPLANTAAUTOPARTISTA = True Else oEmpleado.BoExtension.ADICIONALPLANTAAUTOPARTISTA = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalMantenimientoEdilicio = Trim(HojaExcel.ActiveSheet.Cells(R, 10).Value)
				If xAdicionalMantenimientoEdilicio <> "" Then
					If xAdicionalMantenimientoEdilicio = "SI" Then oEmpleado.BoExtension.AdicionalMantEdilicio = True Else oEmpleado.BoExtension.AdicionalMantEdilicio = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalEmpresa = Trim(HojaExcel.ActiveSheet.Cells(R, 11).Value)
				If xAdicionalEmpresa <> "" Then
					If IsNumeric(xAdicionalEmpresa) Then oEmpleado.BoExtension.AdicionalEmpresa = cDbl(xAdicionalEmpresa) End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalCoordinador = Trim(HojaExcel.ActiveSheet.Cells(R, 12).Value)
				If xAdicionalCoordinador <> "" Then
					If xAdicionalCoordinador = "SI" Then oEmpleado.BoExtension.AdicionalCoordinador = True Else oEmpleado.BoExtension.AdicionalCoordinador = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xTituloTecnico = Trim(HojaExcel.ActiveSheet.Cells(R, 13).Value)
				If xTituloTecnico <> "" Then
					If xTituloTecnico = "SI" Then oEmpleado.BoExtension.Tituo_univ_terc = True Else oEmpleado.BoExtension.Tituo_univ_terc = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xTITULOSECUNDARIO = Trim(HojaExcel.ActiveSheet.Cells(R, 14).Value)
				If xTITULOSECUNDARIO <> "" Then
					If xTITULOSECUNDARIO = "SI" Then oEmpleado.BoExtension.TITULOSECUNDARIO = True Else oEmpleado.BoExtension.TITULOSECUNDARIO = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalTurnicidad = Trim(HojaExcel.ActiveSheet.Cells(R, 15).Value)
				If xAdicionalTurnicidad <> "" Then
					If xAdicionalTurnicidad = "SI" Then oEmpleado.BoExtension.AdicionalTurnicidad = True Else oEmpleado.BoExtension.AdicionalTurnicidad = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalPorTurno = Trim(HojaExcel.ActiveSheet.Cells(R, 16).Value)
				If xAdicionalPorTurno <> "" Then
					If xAdicionalPorTurno = "SI" Then oEmpleado.BoExtension.AdicionalPorTurno = True Else oEmpleado.BoExtension.AdicionalPorTurno = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalCabreada = Trim(HojaExcel.ActiveSheet.Cells(R, 17).Value)
				If xAdicionalCabreada <> "" Then
					If xAdicionalCabreada = "SI" Then oEmpleado.BoExtension.AdicionalCabreada = True Else oEmpleado.BoExtension.AdicionalCabreada = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				xAdicionalEspecialidad = Trim(HojaExcel.ActiveSheet.Cells(R, 18).Value)
				If xAdicionalEspecialidad <> "" Then
					If xAdicionalEspecialidad = "SI" Then oEmpleado.BoExtension.AdicionalEspecialidad = True Else oEmpleado.BoExtension.AdicionalEspecialidad = False End If
				End If
				'-------------------------------------------------------------------------------------------------------'
				If conErrores Then
				    HojaExcel.ActiveSheet.Cells(R, 19).Value = "Empleado No Actualizado (Con Errores)."
					oEmpleado.Workspace.Rollback
				Else
				    If oEmpleado.Workspace.InTransaction Then 
					    oEmpleado.Workspace.Commit
						HojaExcel.ActiveSheet.Cells(R, 19).Value = "Empleado Actualizado Correctamente."
					Else
						HojaExcel.ActiveSheet.Cells(R, 19).Value = "Empleado No Actualizado ."
					End If
				End If
			Else
				HojaExcel.ActiveSheet.Cells(R, 20).Value = "Empleado " & xLegajo & " NO encontrado."
			End If
		End If	
		R = R + 1
	Loop
	
	HojaExcel.ActiveWorkbook.Save
	HojaExcel.Application.Quit
	HojaExcel.Quit
	Call ProgressControlFinish(Self.Workspace)
	If conErrores Then
		MsgBox "El poceso finalizó con errores, verifique la columna (T) del Excel para ver los detalles.", 16, "Error"
	Else
		MsgBox "El poceso finalizó correctamente"
	End If

End SUb
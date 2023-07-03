REM MU - Importar Adicionales desde Excel

'CORREGIDO 
Sub Main ()
	Stop
	xPath = OpenFileDialog("C:\", "Archivo (*.xls)")
	If xPath = "" Then
		MsgBox "No seleccionó ningún archivo. El proceso se cancela.", 48, "Aviso" : Exit Sub
	End If
	' -- EXCEL
	Set HojaExcel = CreateObject("Excel.Application") : HojaExcel.Workbooks.Open xPath : HojaExcel.Sheets("Analista").Select
	R = 2 : xTotalFilas = 0
	Do While Trim(HojaExcel.ActiveSheet.Cells(R, 1).Value) <> ""
	    xTotalFilas = xTotalFilas + 1 : R = R + 1
	Loop
	Call ProgressControl(Self.Workspace, "Importación de Indumentaria desde Excel", 0, xTotalFilas)
	R = 4
	conErrores 	= False
    Set oListaTalleIndumentaria   = ExisteBo(Self,"TipoClasificador","ID","D8CF3B5F-05A7-4D67-A597-809434C3A02A",nil,True,False,"=")
    Set oListaTalleCalzado        = ExisteBo(Self,"TipoClasificador","ID","0BA36B90-5339-48A1-9CFE-3818ADAF84FE",nil,True,False,"=")
	Do While Trim(HojaExcel.ActiveSheet.Cells(R, 1).Value) <> ""
		Call ProgressControlAvance(Self.Workspace, "Fila: " & R & " - Legajo: " & Trim(HojaExcel.ActiveSheet.Cells(R, 1).Value))
		If HojaExcel.ActiveSheet.Cells(R, 12).Value <> "Empleado Actualizado Correctamente." Then 
			xMsgError = "" : xLegajo = "" : xTalleRemera = "" : xTalleBuzo = "" : xTalleCamisa = "": xTalleCampera = "" : xTalleCalzado = "" : xTallePantalon = "" : xTallePrendaCompleta = ""
			xLegajo	= Trim(HojaExcel.ActiveSheet.Cells(R, 1).Value)
            set oEmpleado = Nothing
			Set xViewEmpleado = NewCompoundView(Self,"EMPLEADO", Self.Workspace, nil, True)
			xViewEmpleado.addfilter( NewFilterSpec(NewColumnSpec( "EMPLEADO", "ACTIVESTATUS", "EMPLEADO" ),"=", 0))
			xViewEmpleado.addfilter( NewFilterSpec(NewColumnSpec( "EMPLEADO", "CODIGO", "EMPLEADO" ),"=", xLegajo))
			For Each ee In xViewEmpleado.ViewItems
				Set oEmpleado = ee.bo : Exit For
			Next
			If Not oEmpleado Is Nothing Then
                xTalleCamisa = Trim(HojaExcel.ActiveSheet.Cells(R, 4).Value)
				If xTalleCamisa <> "" Then
                    Set oTalleCamisa = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "NOMBRE", xTalleCamisa, oListaTalleIndumentaria.Valores, True, False, "=")
                    If oTalleCamisa Is Nothing Then
                        conErrores =True
                        LibroExcel.ActiveSheet.Cells(R, 26).Value = chr(13) & "Talle de remera no encontrado. - " & LibroExcel.ActiveSheet.Cells(R, 26).Value 
                    Else
                        set oEmpleado.BoExtension.OBJTALLECAMISA = oTalleCamisa
                    End If
                End If
				'-------------------------------------------------------------------------------------------------------'
                xTallePantalon = Trim(HojaExcel.ActiveSheet.Cells(R, 7).Value)
                If xTallePantalon <> "" Then
                    Set oTallePantalon = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "NOMBRE", xTallePantalon, oListaTalleIndumentaria.Valores, True, False, "=")
                    If oTallePantalon Is Nothing Then
                        conErrores =True
                        LibroExcel.ActiveSheet.Cells(R, 26).Value = chr(13) & "Talle de remera no encontrado. - " & LibroExcel.ActiveSheet.Cells(R, 26).Value 
                    Else
                        set oEmpleado.BoExtension.OBJTALLEPANTALON = oTallePantalon
                    End If
                End If
				'-------------------------------------------------------------------------------------------------------'
                xTalleCalzado = Trim(HojaExcel.ActiveSheet.Cells(R, 10).Value)
                If xTalleCalzado <> "" Then
                    Set oTalleCalzado = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "NOMBRE", xTalleCalzado, oListaTalleCalzado.Valores, True, False, "=")
                    If oTalleCalzado Is Nothing Then
                        conErrores =True
                        LibroExcel.ActiveSheet.Cells(R, 26).Value = chr(13) & "Talle de remera no encontrado. - " & LibroExcel.ActiveSheet.Cells(R, 26).Value 
                    Else
                        set oEmpleado.BoExtension.OBJTALLECALZADO = oTalleCalzado
                    End If
                End If
				'-------------------------------------------------------------------------------------------------------'
                xTalleCampera = Trim(HojaExcel.ActiveSheet.Cells(R, 13).Value)
                If xTalleCampera <> "" Then
                    Set oTalleCampera = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "NOMBRE", xTalleCampera, oListaTalleIndumentaria.Valores, True, False, "=")
                    If oTalleCampera Is Nothing Then
                        conErrores =True
                        LibroExcel.ActiveSheet.Cells(R, 26).Value = chr(13) & "Talle de remera no encontrado. - " & LibroExcel.ActiveSheet.Cells(R, 26).Value 
                    Else
                        set oEmpleado.BoExtension.OBJTALLECAMPERA = oTalleCampera
                    End If
                End If
				'-------------------------------------------------------------------------------------------------------'
				xTalleBuzo = Trim(HojaExcel.ActiveSheet.Cells(R, 16).Value)
				If xTalleBuzo <> "" Then
                    Set oTalleBuzo = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "NOMBRE", xTalleBuzo, oListaTalleIndumentaria.Valores, True, False, "=")
                    If oTalleBuzo Is Nothing Then
                        conErrores =True
                        LibroExcel.ActiveSheet.Cells(R, 26).Value = chr(13) & "Talle de remera no encontrado. - " & LibroExcel.ActiveSheet.Cells(R, 26).Value 
                    Else
                        set oEmpleado.BoExtension.OBJTALLEBUZO = oTalleBuzo
                    End If
                End If
				'-------------------------------------------------------------------------------------------------------'
                xTalleRemera = Trim(HojaExcel.ActiveSheet.Cells(R, 19).Value)
                If xTalleRemera <> "" Then
                    Set oTalleRemera = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "NOMBRE", xTalleRemera, oListaTalleIndumentaria.Valores, True, False, "=")
                    If oTalleRemera Is Nothing Then
                        conErrores =True
                        LibroExcel.ActiveSheet.Cells(R, 26).Value = chr(13) & "Talle de remera no encontrado. - " & LibroExcel.ActiveSheet.Cells(R, 26).Value 
                    Else
                        set oEmpleado.BoExtension.OBJTALLEREMERA = oTalleRemera
                    End If
                End If
				'--------------------------------------------------------------------------------------------------------'
                xTallePrendaCompleta = Trim(HojaExcel.ActiveSheet.Cells(R, 22).Value)
                If xTallePrendaCompleta <> "" Then
                    Set oTallePrendaCompleta = ExisteBo(Self, "ITEMTIPOCLASIFICADOR", "NOMBRE", xTallePrendaCompleta, oListaTalleIndumentaria.Valores, True, False, "=")
                    If oTallePrendaCompleta Is Nothing Then
                        conErrores =True
                        LibroExcel.ActiveSheet.Cells(R, 26).Value = chr(13) & "Talle de remera no encontrado. - " & LibroExcel.ActiveSheet.Cells(R, 26).Value 
                    Else
                        set oEmpleado.BoExtension.OBJTALLEPRENDACOMPLETA = oTallePrendaCompleta
                    End If
                End If
                
			Else
                conErrores = True
				HojaExcel.ActiveSheet.Cells(R, 26).Value = "Empleado " & xLegajo & " NO encontrado."
			End If
            If conErrores Then
				    HojaExcel.ActiveSheet.Cells(R, 26).Value = "Empleado No Actualizado (Con Errores)."
					oEmpleado.Workspace.Rollback
				Else
				    If oEmpleado.Workspace.InTransaction Then 
					    oEmpleado.Workspace.Commit
						HojaExcel.ActiveSheet.Cells(R, 26).Value = "Empleado Actualizado Correctamente."
					Else
						HojaExcel.ActiveSheet.Cells(R, 26).Value = "Empleado No Actualizado ."
					End If
			End If
		End If	
		R = R + 1
	Loop
	
	HojaExcel.ActiveWorkbook.Save
	HojaExcel.Application.Quit
	HojaExcel.Quit
	Call ProgressControlFinish(Self.Workspace)
	If conErrores Then
		MsgBox "El poceso finalizó con errores, verifique la columna (L y/o M) del Excel para ver los detalles.", 16, "Error"
	Else
		MsgBox "El poceso finalizó correctamente"
	End If

End SUb
'importar talles indumentaria

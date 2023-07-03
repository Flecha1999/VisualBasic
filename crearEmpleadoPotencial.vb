REM NEURONAL - EMPLEADO POTENCIAL
Sub Main
   
     
   ok = true
   If aObject.BoExtension.Cuil = "" Then
       AddFailedConstraint "Falta asignar el CUIL del empleado potencial",0
   End If
   If aObject.BoExtension.DNI = "" Then
       AddFailedConstraint "Falta asignar el DNI del empleado potencial",0
	   Exit Sub
   Else
   	   Call CrearPersona (Aobject)
   End If

   If Not aObject.enteasociado Is Nothing Then 
      Set oEmpleado = Nothing
   	  Set oEmpleado = ExisteBo(aObject,"EMPLEADO","ENTEASOCIADO",aObject.enteasociado.ID,nil,True,False,"=")  
   	  If Not oEmpleado Is Nothing Then         
	  	 aObject.codigo   = oEmpleado.codigo
		 aObject.boextension.fechaingreso = oEmpleado.fechaingreso
	     aObject.boextension.FECHAANTIGUEDADRECONOCIDA = oEmpleado.fechaantiguedad    
		 aObject.Empleado = oEmpleado              
      End if
   End if	
   
   If aObject.boextension.estado Is Nothing Then
      AddFailedConstraint "Falta asignar el estado del E.P.",2
	  exit sub
   End If
'  ---------------------------------------------------------------------------------------------------
   
   If aObject.boextension.estado.id = "{E8CFF3B6-83B4-41D3-8663-34A1C6352F30}" Then ' OK SELECCION
   	  If Not PerteneceAGrupo( "SELECCION SELECCION" ) And uCase(NombreUsuario()) <> "DESARROLLO2" Then
   	  	 AddFailedConstraint "Usuario NO autorizado para asignar el estado - OK SELECCION -.", 2
	  Else
      	 If aObject.BoExtension.TipoContratacion Is Nothing Then AddFailedConstraint "Falta asignar el TIPO DE CONTRATACION del empleado potencial", 2 End If
      	 If aObject.BoExtension.Categoria        Is Nothing Then AddFailedConstraint "Falta asignar la CATEGORIA del empleado potencial", 2 End If
      	 If aObject.BoExtension.CentroCostos     Is Nothing Then AddFailedConstraint "Falta asignar el CENTRO DE COSTOS del empleado potencial", 2 End If
      	 If aObject.BoExtension.FechaIngreso     =  EMPTY   Then AddFailedConstraint "Falta asignar la FECHA DE INGRESO del empleado potencial", 2 End If
      	 If aObject.BoExtension.Convenio         Is Nothing Then 
         	AddFailedConstraint "Falta asignar el CONVENIO del empleado potencial", 2 
      	 Else
         	If aObject.BoExtension.Convenio.Id = "{8E6C3E23-2EA1-47C9-808D-14DD8E591718}" Or aObject.BoExtension.Convenio.Id = "{A9714C75-98F1-40EE-B9C4-075A5916587E}" Then
               If aObject.BoExtension.Basico <= 0 Then
               	  AddFailedConstraint "Falta asignar el BÁSICO del empleado potencial", 2 
               End If
         	End If
      	 End If
		 If not aObject.BoExtension.Centrocostos is nothing Then
    	 	If not aObject.BoExtension.Categoria is nothing Then
        	   If aObject.BoExtension.Centrocostos.id = "{5471FBEE-8335-4F68-9F33-6A8DE8ED850B}" Then ' GM JANITORIAL
               	  If aObject.BoExtension.Categoria.id <> "{8A0666DC-CB85-4103-823A-CCB9E70D39C0}" Or aobject.boextension.categoria.id <> "{79244C38-DC25-43A5-9746-FD11260F5017}" Then ' SORByLR - Operario Automotriz Complejidad Técnica O SORByLR - Operario Automotriz Areas Comunes
                  	' AddFailedConstraint "La categoria asignada no es correcta para el centro de costos",2   
				  End If
        	   End If
    		Else 
        	   Addfailedconstraint "Categoria no asignada",2
    		End If
		 Else
    	    Addfailedconstraint "Centro de costos no asignado",2
		 End If		 
	  End If
   End If
   If aObject.boextension.estado.id = "{1103075A-AE6C-42E1-9F99-DA56E9481499}" Then ' OK ADMINISTRACION
   	  If Not PerteneceAGrupo( "SELECCION SELECCION" ) And Not PerteneceAGrupo( "SELECCION ADMINISTRACION" ) And Not PerteneceAGrupo( "CALIPSO RH LIQUIDACIONES" ) And uCase(NombreUsuario()) <> "DESARROLLO2"  And uCase(NombreUsuario()) <> "MVILLARRUEL" And uCase(NombreUsuario()) <> "DESARROLLO3" And uCase(NombreUsuario()) <> "DESARROLLO4" Then
   	  	 AddFailedConstraint "Usuario NO autorizado para asignar el estado - OK ADMINISTRACION -.", 2
	  Else
	     If aObject.BoExtension.ObraSocial      Is Nothing Then AddFailedConstraint "Falta asignar la OBRA SOCIAL del empleado potencial", 2 End If
      	 If aObject.BoExtension.Nacionalidad    Is Nothing Then AddFailedConstraint "Falta asignar la NACIONALIDAD del empleado potencial", 2 End If
      	 'If aObject.BoExtension.Condicion       Is Nothing Then AddFailedConstraint "Falta asignar la CONDICION del empleado potencial", 2 End If
      	 If aObject.BoExtension.FechaNacimiento = EMPTY    Then AddFailedConstraint "Falta asignar la FECHA DE NACIMIENTO del empleado potencial", 2 End If
	  End If
   End If
'  ---------------------------------------------------------------------------------------------------
   If Not aObject.boextension.convenio Is Nothing Then
       'S.O.M., COMERCIO y PARQUES Y JARDINES
      If aObject.boextension.convenio.id = "{CF65CADD-83AA-43BF-AED4-CF7071A2A061}" or _   
	     aObject.boextension.convenio.id = "{D86EE5BA-6F25-4525-9803-5F4FEAA852C9}" or _  
		 aObject.boextension.convenio.id = "{68344A7D-4055-4A30-8881-71ADB7EC707F}" Then
	     
		 aObject.boextension.sindicato = Nothing
	  End if
   End If	  
   
   if not aObject.boextension.categoria is nothing then
      'Fuera de Convenio o Pasantía
	  if aObject.boextension.categoria.id <>"{0D28EF23-AFFC-408B-9D04-8A6A3E43DA69}" and aObject.boextension.categoria.id <> "{65DB7735-23BD-4C39-AF0D-53DC3D01FE59}"  then   
	     aObject.boextension.basico = 0
	  End If
   Else
	  aObject.boextension.basico = 0   
   End if     
	  	    
   
   If aObject.boextension.estado.id = "{E8CFF3B6-83B4-41D3-8663-34A1C6352F30}" or _
	  aObject.boextension.estado.id = "{1103075A-AE6C-42E1-9F99-DA56E9481499}"  Then 'OK Seleccion / OK Administracion
   
      if aobject.BOEXTENSION.puesto is nothing Then 
         AddFailedConstraint "Falta seleccionar el puesto",2   
	  	 ok = false
      End if
   	  if aobject.BOEXTENSION.CENTROCOSTOS is nothing Then 
         AddFailedConstraint "Falta seleccionar el centro de costos",2   
	  	 ok = false
      End if
   	  if aobject.BOEXTENSION.SECTORACTUAL is nothing Then 
      	 AddFailedConstraint "Falta seleccionar la sucursal",2   
	  	 ok = false
      End if      
   
   	  If ok then ' FALTA PREGUNTAR POR EL ESTADO DEL EMPLEADO, SOLO ASIGNAR AL PEDIDO SI EL EMPLEADO ESTA OK

         If Not aObject.boextension.PEDIDO Is Nothing Then
	  
         	xstring2 = StringConexion( "CALIPSO", aobject.WorkSpace ) 
   	  	 	set xcone2 = createobject("adodb.connection")
   	  		xcone2.connectionstring = xstring2
   	  	 	xcone2.connectiontimeout=150
		 
		 	set xpro2 = RecordSet(xCone2, "select top 1 * from producto")
		 	xpro2.close
   		 	xpro2.activeconnection.commandtimeout=0
   		 	xpro2.source="select top 1 * from producto" ' en q2 esta la consulta
   		 	xpro2.open
		 	xpro2.close
      	 	xpro2.activeconnection.commandtimeout=0
		 
		 	xpro2.source= " SELECT UD.PEDIDO_ID, COUNT(*) CANT  "&CHR(13)&"" & _
		 			   " FROM UD_EMPLEADOPOTENCIAL UD WITH(NOLOCK)  "&CHR(13)&"" & _
					   " WHERE ESTADO_ID IN ('E8CFF3B6-83B4-41D3-8663-34A1C6352F30', '3B87B7F9-39AA-44AA-AFFC-A07A966D288C','1103075A-AE6C-42E1-9F99-DA56E9481499','18B17DA1-0A61-4D8B-8752-5BBDCD9F637B')  "&CHR(13)&"" & _
					   " AND UD.PEDIDO_ID = '"&aObject.boextension.PEDIDO.ID&"' "&CHR(13)&"" & _
					   " AND UD.ID <> '"&aObject.boextension.ID&"' "&CHR(13)&"" & _
					   " GROUP BY UD.PEDIDO_ID "
		 	xpro2.open
         	CANTASIGNADOS =0	 
		    do while not xpro2.eof
			   CANTASIGNADOS = int(xpro2("CANT").value)
			   xpro2.movenext
		    loop	
			
			aObject.boextension.PEDIDO.PENDIENTES = aObject.boextension.PEDIDO.CANTPERSONAS - CANTASIGNADOS - 1
			
			            xpro2.close
      	 	xpro2.activeconnection.commandtimeout=0
            xpro2.source = " SELECT PS.ID, COUNT(*) CANT  " & _
            " FROM UD_EMPLEADOPOTENCIAL UD WITH(NOLOCK)   " & _
            " LEFT OUTER JOIN UD_PUESTOS P WITH(NOLOCK) ON P.ID = UD.PEDIDO_ID " & _
            " LEFT OUTER JOIN UD_PROCESOSELECCION PS WITH(NOLOCK) ON PS.PUESTOS_ID = P.BO_PLACE_ID " & _
            " WHERE UD.ESTADO_ID IN ('E8CFF3B6-83B4-41D3-8663-34A1C6352F30', '3B87B7F9-39AA-44AA-AFFC-A07A966D288C','1103075A-AE6C-42E1-9F99-DA56E9481499','18B17DA1-0A61-4D8B-8752-5BBDCD9F637B')  " & _
            " AND PS.ID = '"&aObject.boextension.PEDIDO.BO_PLACE.BO_OWNER.ID&"'  " & _
			" AND UD.ID <> '"&aObject.boextension.ID&"' "& _
            " GROUP BY PS.ID  " 
   		 	xpro2.open
         	CANTASIGNADOS =0	 
		    do while not xpro2.eof
			   CANTASIGNADOS = int(xpro2("CANT").value)
			   xpro2.movenext
		    loop	
			
			aObject.boextension.PEDIDO.BO_PLACE.BO_OWNER.PENDIENTES = aObject.boextension.PEDIDO.BO_PLACE.BO_OWNER.CANT_PERSONAS - CANTASIGNADOS - 1
			
   	        If aObject.boextension.PEDIDO.PENDIENTES = 0 Then	
			   aObject.boextension.pedido.BO_PLACE.BO_OWNER.FECHAINGRESOEFECTIVO = aObject.BoExtension.FechaIngreso 'Le cargo la fecha de este último EP al pedido
			   aObject.boextension.pedido.BO_PLACE.BO_OWNER.PENDIENTES = 1 		 ' PARA QUE ENTRE AL NEURONAL DEL PEDIDO      
			   aObject.boextension.pedido.BO_PLACE.BO_OWNER.PENDIENTES = 0 
			End If
							
         End If	
		 
		 If aObject.boextension.estado.id = "{E8CFF3B6-83B4-41D3-8663-34A1C6352F30}" Then 'OK SELECCION
		    call crearPIPRH (aObject)
		 End If	
		 
		 If aObject.codigo = "" Then
		    aObject.codigo = UltimoLegajo(aObject) 
		 End if
		 	  
      End if
   End If

End Sub



Private Function CrearPersona (Aobject)
   If aObject.enteasociado Is Nothing Then
	  If aObject.boextension.apellido ="" or aObject.boextension.nombre = "" Then
	     AddFailedConstraint "Falta completar el apellido y/o nombre",2
		 exit Function   
	  End If
	  cuil = aObject.boextension.cuil
	  CuilSinGuion  = Replace(cuil, "-", "")
	  RTA = verificaCuit(cuil)
	  If RTA <> "CORRECTO" Then
	     AddFailedConstraint "CUIT INCORRECTO",2 
	     Exit Function
	  End If		  
	  
	  Set xPersona = nothing
	  ' Fisica DNI.	 ' Set xPersona = ExisteBo(aObject,"Personafisica","cuit",CuilSinGuion,nil,True,False,"=")
	  Set xView = NewCompoundView(aObject, "PERSONAFISICA", aObject.Workspace, nil, true)
	  xView.NoFlushBuffers = True		' WITH(NOLOCK)
	  xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("NUMERODOCUMETO"), " = ", aObject.BoExtension.DNI))
	  xView.AddFilter(NewFilterSpec(xView.ColumnFromPath("ACTIVESTATUS"), " = ", 0))
	  If xView.ViewItems.Size > 0 Then
		 Set xPersona = xView.ViewItems.First.Current.BO
	  End If
	  
	  If xPersona Is Nothing Then	  
	  	 Set xsistema = ExisteBo(aObject,"appcalipso","id","BE582E8B-30BE-4F7B-9032-ABF58AE718BF",nil,True,False,"=")
	  	 Set xpersona = crearbo("Personafisica",aObject)
	  	 xsistema.personasfisicas.add(xpersona)
	  	 xpersona.cuit = CuilSinGuion
	  End If
	   
      aObject.enteasociado   = xPersona
	  xpersona.apellido 	 = Replace(Replace(aObject.boextension.apellido, chr(10), ""), chr(13), "")
	  xpersona.PrimerNombre  = Replace(Replace(aObject.boextension.nombre, chr(10), ""), chr(13), "")
	  xpersona.tipodocumento    = ExisteBo(aObject,"tipodocumento","codigo","1",nil,True,False,"=") 'D.N.I.
	  xpersona.numerodocumeto   = aObject.boextension.dni
	  xpersona.fechanacimiento  = aObject.boextension.fechanacimiento	
	  xPersona.nacionalidad     = aObject.boextension.nacionalidad	  	  
   End If
End Function



Private Function CrearPIPRH (Aobject)

	If Aobject.boextension.indumentaria.size < 1 Then Exit Function 
	
	If Not Aobject.BOEXTENSION.REFERENCIAPIPRH Is Nothing Then Exit Function
	
	For each xItemEP in Aobject.boextension.indumentaria
	   If xItemEP.producto Is Nothing Then 
	      MsgBox "Existe indumentaria sin producto asignado por lo que no se va a crear el PIPRH" 
		  Exit Function
	   End If	  
	Next
	
	res = MsgBox("¿Desea crear PIP RH?",4)
	if res <> 6 then	
	   Exit Function
	End If
	
	set xUO			 		= InstanciarBO("{84D92774-EB25-4599-8137-0403FA1CAF40}", "UOINVENTARIO", Aobject.Workspace)	
	set xTipoPIP			= InstanciarBO("{963541B8-C7A5-4037-8BB6-2A0CEAC0E096}", "IMPUTACIONCONTABLE", Aobject.Workspace)	 ' Ingreso
	set xTrNueva 			= CrearTransaccion("PIPRH", xUO)
	xTrNueva.imputacioncontable = xTipoPIP
	
	For each xItemEP in Aobject.boextension.indumentaria
	   Set Item = crearBO ("UD_EMPLEADOPIPRH",Aobject)
	   Item.centrocostos  	 = xItemEP.centrocostos
	   Item.EMPLEADO  	 	 = Aobject
	   Item.TIPOINDUMENTARIA = xItemEP.TIPOINDUMENTARIA
	   Item.FECHAENTREGA 	 = xItemEP.FECHAENTREGA
	   Item.talle 			 = xItemEP.talle
	   Item.cantidad 		 = xItemEP.cantidad
	   Item.producto		 = xItemEP.producto
	   Item.descripcion		 = xItemEP.descripcion	   
	   xTrNueva.boextension.items.add(Item)	   
	Next
	set xItemNuevo = CrearItemTransaccion(xTrNueva)
	xItemNuevo.Referencia				= InstanciarBO("11F1EB95-8BAC-474E-A848-716948A2579D", "CONCEPTOCONTABLE", Aobject.Workspace)	' Concepto Generico.
	xItemNuevo.Cantidad.Cantidad		= 1.0
	xItemNuevo.Valor.Importe			= 1.0	
	Aobject.BOEXTENSION.REFERENCIAPIPRH = xTrNueva

End Function



Private Function UltimoLegajo (Aobject)
   
   xstring2 = StringConexion( "CALIPSO", aobject.WorkSpace ) 
   set xcone2 = createobject("adodb.connection")
   xcone2.connectionstring = xstring2
   xcone2.connectiontimeout=150
		 
   set xpro2 = RecordSet(xCone2, "select top 1 * from producto")
   xpro2.close
   xpro2.activeconnection.commandtimeout=0
   xpro2.source="select top 1 * from producto" ' en q2 esta la consulta
   xpro2.open
   xpro2.close
   xpro2.activeconnection.commandtimeout=0
		 
   xpro2.source= " SELECT MAX(cast(Q.CODIGO as int)) ULTIMOLEGAJO"&CHR(13)&"" & _
   				 " FROM ( "&CHR(13)&"" & _
				 " SELECT CODIGO FROM EMPLEADO WITH(NOLOCK) "&CHR(13)&"" & _
				 " WHERE ACTIVESTATUS <> 2 "&CHR(13)&"" & _				 
				 " UNION ALL "&CHR(13)&"" & _
				 " SELECT CODIGO FROM EMPLEADOPOTENCIAL WITH(NOLOCK) WHERE LEN(CODIGO)<6 "&CHR(13)&"" & _
				 ") Q "
   xpro2.open
         		 
   do while not xpro2.eof
	  UltimoLegajo = xpro2("ULTIMOLEGAJO").value + 1
	  xpro2.movenext
   loop	

End Function

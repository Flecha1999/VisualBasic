' Fx Crear Empleado - pruebas 3
Sub Main()
    stop
	If Not Self.boextension.estado Is Nothing Then
	   If Self.boextension.estado.id <> "{1103075A-AE6C-42E1-9F99-DA56E9481499}" Then 'OK ADMINISTRACION
		  MsgBox "Estado incorrecto para crear el Empleado. Debe estar en OK ADMINISTRACION"
		  exit sub     
	   End If
	Else
	   MsgBox "Empleado potencial sin Estado"
	   exit sub  
	End If 	
	

	set xempleado = Nothing
	Set xapp = Self.workspace
    Set EP = self
	If Not EP.empleado Is Nothing Then
	   If EP.empleado.activestatus = 0 Then
	      MsgBox "Ya se creó el emplado correspondiente a este empleado potencial"  
	      EXIT SUB
	   Else
	      Set xempleado = EP.empleado 
	   	  If xempleado.activestatus <> 0 Then
	         set xVector = NewVector()
	   	  	 set xBucket = NewBucket()
	   	  	 xBucket.Value = "UPDATE EMPLEADO SET ACTIVESTATUS = 0 WHERE ID = '" & xempleado.id & "'"
	   	  	 xVector.Add(xBucket)
	   	  	 ExecutarSQL xVector, "DistrObj", "", SELF.workspace, -1
	   	  	 senddebug "------>> Haciendo Update de EMPLEADO -> ACTIVANDOLO"
	   	  	 call WorkSpaceCheck( SELF.WorkSpace )	
	   	  	 Set EP.empleado = xempleado
			 xempleado.boextension.reingreso = True
			 xempleado.boextension.CuentaFuturos = 0
			 xempleado.MOTIVOBAJA = Nothing
			 
	      End If	
	   End If
	End If
	
	Set xsistema		= ExisteBo(Self,"appcalipso","id","BE582E8B-30BE-4F7B-9032-ABF58AE718BF",nil,True,False,"=")
	Set xcompania		= ExisteBo(Self,"compania","id","63E432B5-726E-4195-8F9D-2025C20D7467",nil,True,False,"=")
	Set oTiposNomina	= ExisteBo(Self,"TipoClasificador","nombre","Tipo de Pago",nil,True,False,"=")
	
	xbanco		= ""
	sucursal	= ""
	tipocuenta	= ""
	xcuenta		= ""
	cbu			= ""
	
	Set xpersona = EP.enteasociado

	If EP.boextension.calle <>"" Then
	   Set xDomicilio		    = crearbo("domicilio",Self)	   
	   xpersona.domicilios.add(xDomicilio)
	   Set xTipoDom		        = InstanciarBO( "{D9CA28CC-404D-46DC-A84F-D3B7DD383623}", "TIPODOMICILIO", SELF.Workspace )
	   xDomicilio.TipoDomicilio	= xTipoDom
	   xDomicilio.pais			= InstanciarBO( "{7AD1D8B0-DC00-4DD6-A615-A89577297AAB}", "PAIS", SELF.Workspace )  
	   xDomicilio.provincia		= EP.boextension.provincia
	   xDomicilio.ciudad	    = EP.boextension.localidad
	   xDomicilio.calle			= EP.boextension.calle
	   If IsNumeric(EP.boextension.nro) Then xDomicilio.numero = EP.boextension.nro	   
	   xDomicilio.barrio		= EP.boextension.barrio  
	   xDomicilio.codpos		= EP.boextension.CODIGOPOSTAL
	   xDomicilio.piso			= EP.boextension.piso
	   xDomicilio.puerta		= EP.boextension.depto 
	   xDomicilio.boextension.referenciaDomicilio = EP.boextension.REFERENCIADOMICILIO
	End If
				 
	Set xTELEFONO=Nothing 
			  
	If EP.boextension.TELEFONOPARTICULAR <> "" Then
	   Set xTELEFONO = crearbo("TELEFONO",Self)
	   xpersona.TELEFONOS.add(xTELEFONO)
	   Set XTIPOTEL                = ExisteBo(Self,"TIPOTELEFONO","ID","8C535294-FBEC-433B-B25F-639111DF861C",nil,True,False,"=") 'celular
	   xTELEFONO.TIPOTELEFONO      = XTIPOTEL
	   xTELEFONO.NUMERO            = EP.boextension.TELEFONOPARTICULAR
	   xpersona.telefonoprincipal  = xTELEFONO
	End If
				 	
	If EP.boextension.TELEFONOALTERNATIVO <> "" Then
	   Set xTELEFONO1				 = crearbo("TELEFONO",Self)
	   xpersona.TELEFONOS.add(xTELEFONO1)				
	   Set XTIPOTEL				 = ExisteBo(Self,"TIPOTELEFONO","ID","2CE3E93B-385C-4443-AE27-E4356F2A2E6E",nil,True,False,"=")'particular
	   xTELEFONO1.TIPOTELEFONO	 = XTIPOTEL
	   xTELEFONO1.NUMERO		 = EP.boextension.TELEFONOALTERNATIVO  
	End If				 
	    
	If xempleado is Nothing Then
	   Set xempleado = crearbo("empleado",Self)
	   xcompania.empleados.add(xempleado)
	   xempleado.enteasociado = xpersona
	   xempleado.codigo	   = EP.codigo ' ObternerUltimoLegajo(self.workspace)
	   xempleado.unidadoperativa = Self.unidadoperativa
	   xpersona.fechanacimiento  = EP.boextension.fechanacimiento	
	   xPersona.nacionalidad     = EP.boextension.nacionalidad	
	End IF
	   
	If EP.boextension.categoria is Nothing Then
	   MSGBOX"Falta Categoria"
    End If
			 
	xempleado.categoria = EP.boextension.categoria  
	xempleado.basico_importe= EP.boextension.BASICO
	xempleado.SITREVISTA1="01"
	xempleado.DIASITREVISTA1="01"
	xempleado.situacion=EP.boextension.situacion
	
	' 18/01/2023 Agregada antiguedad vacaciones PS
	xempleado.boextension.antiguedadvacaciones = EP.boextension.antiguedadvacaciones

	
	If xempleado.boextension is Nothing Then
	   Set oboextension=crearbo("ud_empleado",Self)
	   oboextension.bo_owner = xempleado
	   xempleado.boextension = oboextension
	End If

	If EP.boextension.condicion Is Nothing Then
	   xempleado.condicionempleado = InstanciarBO( "817AC20C-0BA7-4D59-9B1E-E9D3929A013C", "CONDICION", self.Workspace )
	Else
	   xempleado.condicionempleado = EP.boextension.condicion
    End If
	
	If Not EP.boextension.TALLECAMISA1   Is Nothing Then xempleado.boextension.casaca   = EP.boextension.TALLECAMISA1.CODIGO
	If Not EP.boextension.TALLEPANTALON1 Is Nothing Then xempleado.boextension.pantalon = EP.boextension.TALLEPANTALON1.CODIGO
	If Not EP.boextension.TALLECALZADO1 Is Nothing Then xempleado.boextension.calzado   = EP.boextension.TALLECALZADO1.CODIGO
		   
	xempleado.boextension.PRIMARIO      = EP.boextension.PRIMARIO
	xempleado.boextension.SECUNDARIO    = EP.boextension.SECUNDARIO
	xempleado.boextension.TERCIARIO     = EP.boextension.TERCIARIO
	xempleado.boextension.UNIVERSITARIO = EP.boextension.UNIVERSITARIO
	'xempleado.boextension.formacion=xformacion
			 
	If EP.boextension.PORCENTAJEJORNADA<>100 Then
	   xempleado.boextension.porcentajejornada=EP.boextension.PORCENTAJEJORNADA
	End If

	xempleado.boextension.adicional28=EP.boextension.ART28GESTAMP			  
	xempleado.sector=EP.boextension.sectoractual
		
	If not EP.boextension.centrocostos is Nothing Then
	   xempleado.centrocostos = EP.boextension.centrocostos
	else
	   MSGBOX "Falta el Centro de Costos"
	End If
	
	xempleado.fechaingreso                     = EP.boextension.fechaingreso
	xempleado.fechaantiguedad                  = EP.boextension.FECHAANTIGUEDADRECONOCIDA
	xempleado.boextension.AntiguedadVacaciones = EP.boextension.FECHAANTIGUEDADRECONOCIDA
					  
	Set banco=ExisteBo(Self,"banco","enteasociadosucursal",xbanco,NIL,True,False,"=")
	xempleado.banco=banco
	If not banco is Nothing Then
	   for each xb in banco.tiposcuenta
	      If xb.nombre=tipocuenta Then
	         xempleado.tipocuenta=xb
	   		 exit for
	      End If
	   next
	End If
	
	xempleado.cbu=cbu
	xempleado.numerocuenta=xcuenta
	xempleado.boextension.sucursalBANCO=sucursal
	xempleado.tipocontratacion=EP.boextension.tipocontratacion
	xempleado.perfil=EP.boextension.CONVENIO					  
	If xempleado.perfil is Nothing Then
	   MsgBox "Falta Perfil/Convenio"
	End If

    xempleado.BoExtension.TipoNomina = EP.boextension.TIPONOMINA
	xempleado.puesto=EP.boextension.puesto
	xEmpleado.enteasociado.sexo = EP.boextension.sexo
	xEmpleado.enteasociado.estadocivil = EP.boextension.estadocivil				
	xEmpleado.boextension.condicionGanancias = InstanciarBO("B1B97578-CC93-4F38-A11C-0AC162B89851", "ItemTipoClasificador", Self.workspace) 'Grupo 0
	xEmpleado.boextension.liquidaGanancias   = True				
	xempleado.ZONA = EP.boextension.zona
	xempleado.obrasocial = EP.boextension.OBRASOCIAL

	If not xempleado.perfil is Nothing Then
	   If xempleado.Perfil.Id = "{8A8D8A26-03BF-422D-86E4-69F6E0E0DF62}" Then ' Es SORBYL
          xMensajeSorbyl = " El empleado "&xEmpleado.Descripcion&" pertenece al convenio SORBYL. Recordar Dar de Alta en el Sindicato"
          Call MsgBox (xMensajeSorbyl, 64, "Información")
		  xempleado.BoExtension.Validar = False
	   End If
	
	   ' ----- Convencionado ---- '
	   If xempleado.perfil.id = "{8E6C3E23-2EA1-47C9-808D-14DD8E591718}" Then ' Fuera de Convenio
	      xempleado.convencionado = False
	   Else
		  xempleado.convencionado = True
	   End If
	End If
				
	xempleado.boextension.AdicionalPlantaAuto = EP.boextension.ADICIONALPLANTA
	xempleado.sindicato=EP.boextension.SINDICATO
	
	'Indumentaria	
	For Each xItem in EP.boextension.indumentaria
   	   set xTalle = CrearBO("UD_EMPLEADOTALLES", xempleado)
	   set xTalle.Producto		= xItem.Producto
	   xTalle.Descripcion		= xItem.Descripcion
	   xTalle.Cantidad			= xItem.Cantidad
	   xempleado.BOExtension.IndumentariaTalles.Add(xTalle)	   
	Next

	EP.empleado = xempleado
	EP.boextension.FECHADISPONIBILIDAD = Date()
	
	'cambio el EP por el Empleado en el PIPRH
	If Not EP.BOEXTENSION.REFERENCIAPIPRH Is Nothing Then
	   For each xItem in EP.BOEXTENSION.REFERENCIAPIPRH.boextension.items
	      If Not xItem.empleado Is Nothing Then
		     If xItem.empleado.id = Ep.id Then
			    xItem.empleado = xEmpleado
			 End If
		  End If
	   Next 
	End If
		 
	a = workspacecheck(Self.workspace)
	If a = 0 Then
	   ok="Procesado - " '& HojaExcel.ActiveSheet.Cells(i, 54).Value
   	   Self.boextension.estado = InstanciarBO( "{18B17DA1-0A61-4D8B-8752-5BBDCD9F637B}", "ITEMTIPOCLASIFICADOR", SELF.Workspace ) 'Empleado Creado
	   call WorkSpaceCheck( self.WorkSpace )
    else
	   ok="NO Procesado"
	End If 
End Sub

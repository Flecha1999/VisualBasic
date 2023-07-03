REM BTN - EMPLEADO POTENCIAL - Crear Empleados
Sub Main
   stop
   'If Not PerteneceAGrupo( "CALIPSO RH LIQUIDACIONES" ) Then
   '    Call MsgBox ("Usuario no autorizado.",64,"Información")
   '    Exit Sub
   ' End If
   cont = 0
   contN = 0
   empleadosCreados =""
   empleadosNOCreados =""
   
   For Each oRecord In Container        
      If Container.Size >= 1 Then
	     Set self = oRecord
		 
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
    	 Set EP   = self

		 If Not EP.empleado Is Nothing Then
		 	if year(EP.boextension.fechaingreso) <> year(date) Then
				MSGBOX "La fecha de ingreso de: " & EP.boextension.NOMBRE & "DNI: "& EP.boextension.DNI & "es distinto al año en curso y no es posible asignar esa fecha"
			exit sub
		 	End If
		 	if year(xempleado.fechaantiguedad ) <> year(date) Then
				MSGBOX "La fecha de antiguedad de: " & EP.boextension.NOMBRE & "DNI: "& EP.boextension.DNI & "es distinto al año en curso y no es posible asignar esa fecha"
			exit sub
		 	End If
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
	   		xDomicilio.pais				= InstanciarBO( "{7AD1D8B0-DC00-4DD6-A615-A89577297AAB}", "PAIS", SELF.Workspace )  
	   		xDomicilio.provincia		= EP.boextension.provincia
	   		xDomicilio.ciudad			= EP.boextension.localidad
	   		xDomicilio.calle			= EP.boextension.calle
			xPersona.domicilioprincipal = xDomicilio
	   		If IsNumeric(EP.boextension.nro) Then xDomicilio.numero = EP.boextension.nro	   
	   		   xDomicilio.barrio		= EP.boextension.barrio  
	   		   xDomicilio.codpos		= EP.boextension.CODIGOPOSTAL
	   		   xDomicilio.piso			= EP.boextension.piso
	   		   xDomicilio.oficina		= EP.boextension.depto 
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
			
			If EP.BoExtension.EMAIL <> "" Then
                Set oEmail 		  = CrearBO("direccionelectronica", Self)
                Set oTipoDir	  = ExisteBO(Self, "TIPODIRECCIONELECTRONICA", "ID", "{0DF8B0AA-423D-4A85-854F-D2507BB4687B}", nil, true, false, "=")
                oEmail.TIPODIRECCIONELECTRONICA    = oTipoDir
                xPersona.direccioneselectronicas.Add(oEmail)
                oEmail.DireccionElectronica        = EP.boextension.EMAIL
                xPersona.DirecElectronicaprincipal = oEmail
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
            If Not EP.boextension.TALLECAMISA1   Is Nothing Then xempleado.boextension.casaca   = EP.boextension.TALLECAMISA1.CODIGO
			If Not EP.boextension.TALLEPANTALON1 Is Nothing Then xempleado.boextension.pantalon = EP.boextension.TALLEPANTALON1.CODIGO
			If Not EP.boextension.TALLECALZADO1 Is Nothing Then xempleado.boextension.calzado   = EP.boextension.TALLECALZADO1.CODIGO
               
			hoy = day(date)
			mes = month(date)
			NombreMesNumero(mes)
        		numeroSemanaMes= (hoy - 1) \ 7 + 1
			If numeroSemanaMes = 1 Then
        		semanaMes = "1ra Semana del mes " & NombreMesNumero(mes) 
			End If
			If numeroSemanaMes = 2 Then
        		semanaMes = "2da Semana del mes " & NombreMesNumero(mes) 
			End If
			If numeroSemanaMes = 3 Then
        		semanaMes = "3ra Semana del mes " & NombreMesNumero(mes) 
			End If
			If numeroSemanaMes = 4 Then
        		semanaMes = "4ta Semana del mes " & NombreMesNumero(mes) 
			End If

                  set xVisualVar = VisualVarEditor("Codigo identificador de importacion de empleado")
		      call AddVarString(xVisualVar, "00CODIGO", "Cód. de imp. Empleado", "Ingrese",semanaMes )
		      aceptar = ShowVisualVar(xVisualVar)
		      cod = Trim(GetValueVisualVar(xVisualVar, "00CODIGO", "Ingrese"))
			xempleado.boextension.IdentificadorEmplmportacion = cod
		 
  			If Not EP.boextension.OBJTALLEREMERA Is Nothing Then xempleado.boextension.OBJTALLEREMERA = EP.boextension.OBJTALLEREMERA
    		If Not EP.boextension.OBJTALLECAMPERA Is Nothing Then xempleado.boextension.OBJTALLECAMPERA = EP.boextension.OBJTALLECAMPERA
    		If Not EP.boextension.OBJTALLECAMISA Is Nothing Then xempleado.boextension.OBJTALLECAMISA = EP.boextension.OBJTALLECAMISA
			If Not EP.boextension.OBJTALLECALZADO Is Nothing Then xempleado.boextension.OBJTALLECALZADO = EP.boextension.OBJTALLECALZADO
    		If Not EP.boextension.OBJTALLEBUZO Is Nothing Then xempleado.boextension.OBJTALLEBUZO = EP.boextension.OBJTALLEBUZO
    		If Not EP.boextension.OBJTALLEPRENDACOMPLETA Is Nothing Then xempleado.boextension.OBJTALLEPRENDACOMPLETA = EP.boextension.OBJTALLEPRENDACOMPLETA
    		If Not EP.boextension.OBJTALLEPANTALON Is Nothing Then xempleado.boextension.OBJTALLEPANTALON = EP.boextension.OBJTALLEPANTALON
		   
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
	        Else
	   		   MSGBOX "Falta el Centro de Costos"
			End If
			xempleado.fechaingreso                     = EP.boextension.fechaingreso
			xempleado.fechaantiguedad                  = EP.boextension.FECHAANTIGUEDADRECONOCIDA
			xempleado.boextension.AntiguedadVacaciones = EP.boextension.AntiguedadVacaciones 
					  
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
			   
			   ' ----- Tipo Nomina Jornalizado ---- '
			   If xEmpleado.Perfil.Id = "{75D7224D-06AF-423F-802C-76B86430179F}" Or xempleado.perfil.id = "{BDD8FB00-8B97-4C7E-A45A-C637543598B3}" Then	' SMATA / U.O.M.
			   	  xempleado.BoExtension.TipoNomina = ExisteBO(Self, "ITEMTIPOCLASIFICADOR", "ID", "{520C6DA4-1EE7-40BD-8E21-8E0540B86E43}", nil, true, false, "=") ' JORNALIZADO
			   Else
				  xempleado.BoExtension.TipoNomina = ExisteBO(Self, "ITEMTIPOCLASIFICADOR", "ID", "{2296A906-602E-4821-B14C-C66522679879}", nil, true, false, "=") ' MENSUALIZADO
			   End If
	        End If
				
			xempleado.boextension.AdicionalPlantaAuto = EP.boextension.ADICIONALPLANTA
			' En el caso de UOM no se hereda el sindicato del empleado potencial
			If xEmpleado.Perfil.Id <> "{BDD8FB00-8B97-4C7E-A45A-C637543598B3}" Then ' UOM
		  	   	xEmpleado.Sindicato = EP.boextension.SINDICATO
	        End If
			
			If xEmpleado.perfil.id = "{8A8D8A26-03BF-422D-86E4-69F6E0E0DF62}" Then 'sorb
			    xEmpleado.Sindicato = ExisteBo(xEmpleado,"sindicato","id","09F3E85E-8ACF-41FC-8DAE-7804656E132C", nill, True, False,"=")
			ElseIf xempleado.perfil.id = "{495CED3E-3998-44B1-88B3-1785BBDC93BC}" Then 'soelsac
				xEmpleado.Sindicato = ExisteBo(xEmpleado,"sindicato","id","68B79CC4-6186-49B3-AE28-7CA23233D1DC", nill, True, False,"=")
			End If
			
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

'			Examenes Médicos
			For Each xItem In EP.BoExtension.ExamenesMedicos
				Set oExMed = CrearBO("UD_EXAMENESMEDICOS", xEmpleado)
				Set oExMed.TIPOEXAMEN    = xItem.TIPOEXAMEN
				Set oExMed.RESULTADO     = xItem.RESULTADO
				oExMed.OBSERVACION       = xItem.OBSERVACION
				oExMed.RECEPCIONESTUDIOS = xItem.RECEPCIONESTUDIOS
				Set oExMed.CENTRO_MEDICO = xItem.CENTRO_MEDICO
				oExMed.FECHAEXAMEN 		 = xItem.FECHAEXAMEN
				oExMed.ADJUNTO			 = xItem.ADJUNTO
				If Not xItem.MOTIVOEXAMEN Is Nothing Then Set oExMed.MOTIVOEXAMEN = xItem.MOTIVOEXAMEN
				xEmpleado.BoExtension.ExamenesMedicos.Add(oExMed)
			Next
			
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
		 If a = 0 And Then
		    cont = cont + 1
			empleadosCreados = empleadosCreados & xEmpleado.descripcion & Chr(13)
   	   		Self.boextension.estado = InstanciarBO( "{18B17DA1-0A61-4D8B-8752-5BBDCD9F637B}", "ITEMTIPOCLASIFICADOR", SELF.Workspace ) 'Empleado Creado
	   		call WorkSpaceCheck( self.WorkSpace )
         else
		    contN = contN + 1
			empleadosNOCreados = empleadosNOCreados & xEmpleado.descripcion & Chr(13)
	     End If 
      End If
   Next
   
   If contN > 0 Then
      MsgBox "CREADOS: "& cont & chr(13) & empleadosCreados & Chr(13) & Chr(13) & "NO CREADOS: "&contN&":" &chr(13) & empleadosNOCreados
   Else
      MsgBox "CREADOS: "& cont & chr(13) & empleadosCreados
   End If	  
End Sub

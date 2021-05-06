'' @script EnviaAfip 3.6.3   
'06/05/2021 Se Agrego Opcional para WSMTXCA 
'23-04-2021 Se agregó para todos que los RI Tierra del Fuego que van por Factura A se tomen como EX 

'EnviaWsfe  Con WSFE1 y WSMTXCA
'descripcion Autoriza el comprobante conectandose al servicio web de la AFIP con detalle de Productos.
'-------------------------------------------------------------------------------
Public AlicuotaTributo,	BaseImponibleTributo,DescripcionTributo,ImporteTributo,aToken,aSign,nro_calipso_Actual,id_Solicitud,ExisteSolicitud,sCbu,sAlias,_
	   nContribuyente,imp_Tolerancia,cPesos,cDolares,cEuros,cReales,xValid,xSintecrom,xZSA,xCabal,xEgramar,xProdalsa,xStenfar,xCilo,xBinova,xWSMTXCA,xUtilizaPathInstalacionSA,CAE,_
           FuerzaDolarizacionComprobante,moneda_id,CotizacionDolar


Private gInteractivo


Sub Main()
Stop
    
'*********************************************************************************************************************************************
	'Variables Modificables segun el Cliente
'*********************************************************************************************************************************************
	'Tolerancia de Centavos en Total del Comproabnte
	imp_Tolerancia 	= 0.02  
	FuerzaDolarizacionComprobante = False
	CotizacionDolar = 1
	

	set pComprobante = aComprobante.Value
	gInteractivo	 = aInteractivo.Value
	


'Path del Certificado
'Si le asignamos True, usa el Path de instalación de Sistemas Agiles. Ojo!!! El problema de esto es el mantenimiento, porque hay que modificarlo en todas las maquinas
'Si le asignamos False, utiliza directamente los que están en la red y que tenemos que asignar en los parametros "FACTURAELECTRONICA_CERTIFICADO" y "FACTURAELECTRONICA_CLAVE" con el path completo. Ejemplo: Z:\FE\calipso_24e52a790ad63ee0.crt o \\SQLSERVER\Util\fe\calipso---.crt 
xUtilizaPathInstalacionSA = False
	
'Empresa
xValid     = False
xSintecrom = False
xZSA       = True
xCabal     = False
xEgramar   = False
xProdalsa  = False
xBinova    = False
xStenfar   = False
xCilo	  = False				  

'Monedas (Comentar las que no corresponden)
	

'**Valid**
If xValid Then
	cPesos 		= "{84C1C4E2-F92E-4B1C-9C60-7E6D56ADC745}"
	cDolares 	= "{B16A4F9A-A3FA-4D35-A08C-0225463FB393}"
	cReales		= "{}"
	CEuros		= "{}"	

	'Si va a utilizar WSMTXCA poner True, si va a Utilizar WSFE1 Poner False 
	xWSMTXCA = False  
End If

'**Sintecrom 
If xSintecrom Then
	cPesos 			= "{77C7D924-9BC2-416E-A379-100ABCA96050}"
	cDolares 		= "{D38FF449-A3B1-4720-BC1B-6218698E4474}"
	cReales			= "{}"
	CEuros			= "{}"

	'Si va a utilizar WSMTXCA poner True, si va a Utilizar WSFE1 Poner False 
	xWSMTXCA = False  
End If

'**Cabal 
If xCabal Then
	cPesos 			= "{1DCB0F1F-4183-47D7-A1AA-0D112AAC5B37}"
	cDolares 		= "{032CF9F6-D838-4312-8D06-B6EFD59042FD}"
	cReales			= "{797F3B86-2ECB-4598-96DB-0563A489DBF1}"
	CEuros			= "{}"

	'Si va a utilizar WSMTXCA poner True, si va a Utilizar WSFE1 Poner False 
	xWSMTXCA = False  
End If

'**ZSA 
If xZSA Then
	cPesos 			= "{9B2CCAA7-F5B5-4C23-85CA-8526537BBF0B}"
	cDolares 		= "{59468A63-4488-4BC1-915E-83D85F88DE81}"
	cReales			= "{71B7C89A-36D3-4EFB-89DD-510F64984F02}"
	CEuros			= "{CD44DC9A-2AC4-41DC-A5E6-3494BA272D83}"

	'Si va a utilizar WSMTXCA poner True, si va a Utilizar WSFE1 Poner False 
	xWSMTXCA = False  


	If pComprobante.BOExtension.DOLAR_TOTAL > 0 Then
	   If (pComprobante.BOExtension.MonedaZSA.Id = cDolares) Then
		  moneda_id     = "DOL"
	   ElseIf (pComprobante.Total.UnidadValorizacion.Id = cReales) Then
		  moneda_id     = "012"
	   ElseIf (pComprobante.Total.UnidadValorizacion.Id = cEuros) Then
		  moneda_id     = "060"
	   End If
	   FuerzaDolarizacionComprobante = True
	   CotizacionDolar = pComprobante.BOExtension.CotizacionZSA
	End If

End If
	
'**Egramar 
If xEgramar Then
	cPesos 			= "{76C69765-3DAE-11D5-B059-004854841C8A}"
	cDolares 		= "{76C69768-3DAE-11D5-B059-004854841C8A}"
	cReales			= "{20F22EFB-3895-4B43-A055-3A86A0CD9D5D}"
	CEuros			= "{03ED5D02-3947-4CB0-B27B-629CAFAE99E2}"

	'Si va a utilizar WSMTXCA poner True, si va a Utilizar WSFE1 Poner False 
	xWSMTXCA = False  
End If
	
'**Prodalsa 
If xProdalsa Then
	cPesos 			= "{9B2CCAA7-F5B5-4C23-85CA-8526537BBF0B}"
	cDolares 		= "{59468A63-4488-4BC1-915E-83D85F88DE81}"
	cReales			= "{}"
	CEuros			= "{}"

	'Si va a utilizar WSMTXCA poner True, si va a Utilizar WSFE1 Poner False 
	xWSMTXCA = True  
End If

'**Binova**
If xBinova Then
	cPesos 			= "{9EDF155B-C8FF-4273-8922-8E9281028137}"
	cDolares 		= "{F3B7E2B1-1D09-406A-BA68-87E46B38C29A}"
	cReales			= "{}"
	CEuros			= "{}"	

	'Si va a utilizar WSMTXCA poner True, si va a Utilizar WSFE1 Poner False 
	xWSMTXCA = False  
End If
If xStenfar Then
'**Stenfar 
	cPesos 			= "{13207D68-8854-11D5-B08A-004854841C8A}"
	cDolares 		= "{13207D6B-8854-11D5-B08A-004854841C8A}"
	cReales			= "{D8C0F91C-49C0-463C-A68D-26CD5FB7DC05}"
	CEuros			= "{D783CBDE-FA11-4B0A-8590-0018AF050E3C}"
	xWSMTXCA = False
End if		
If xCilo Then
	cPesos 			= "{76C69765-3DAE-11D5-B059-004854841C8A}" :	cDolares 		= "{76C69768-3DAE-11D5-B059-004854841C8A}"
	cReales			= "{}" :	CEuros			= "{}"	
	xWSMTXCA = False
End if

	
'*********************************************************************************************************************************************
'*********************************************************************************************************************************************	
	aToken			= ""
	aSign			= ""
	
	ausuario		= nombreusuario()
	aComputername	= GetComputerName(  )
	aNuevoTicketAccesso = True
	
'Obtengo Parametros 	
	Certificado  		= ExisteBO( pComprobante, "ASOCIACION", "Clave", "FACTURAELECTRONICA_CERTIFICADO", nil, FALSE, FALSE, "=" ).valor
	ClavePrivada    	= ExisteBO( pComprobante, "ASOCIACION", "Clave", "FACTURAELECTRONICA_CLAVE", nil, FALSE, FALSE, "=" ).valor
	xPathXml			= ExisteBO( pComprobante, "ASOCIACION", "Clave", "FACTURAELECTRONICA_PATHLOG", nil, FALSE, FALSE, "=" ).valor
	nContribuyente	= ExisteBO( pComprobante, "ASOCIACION", "Clave", "FACTURAELECTRONICA_CUIT", nil, FALSE, FALSE, "=" ).valor
	sCbu				= ExisteBO( pComprobante, "ASOCIACION", "Clave", "MIPYMES_CBU", nil, FALSE, FALSE, "=" ).valor
	sAlias 			= ExisteBO( pComprobante, "ASOCIACION", "Clave", "MIPYMES_ALIAS", nil, FALSE, FALSE, "=" ).valor
	
	If xWSMTXCA Then
       sServicio = "wsmtxca"
    Else
       sServicio = "wsfe"
    End If
	
	xTicketAcceso	= False
	
	If xWSMTXCA Then
	   'xTicketAcceso = ObtenerTicketAccessoWSMTXCA (pComprobante,nContribuyente,sServicio,aComputername,ausuario,Certificado,ClavePrivada)
	   xTicketAcceso = ObtenerTicketAccesso (pComprobante,nContribuyente,sServicio,aComputername,ausuario,Certificado,ClavePrivada)
    Else
	   xTicketAcceso = ObtenerTicketAccesso (pComprobante,nContribuyente,sServicio,aComputername,ausuario,Certificado,ClavePrivada)
    End If
	
'30  Crear objeto interface Web Service de Factura Electrónica de Mercado Interno

	If xTicketAcceso Then 
		
		set WSFACT		= CreaWSFACT(aToken,aSign,nContribuyente)
   
'40 Establezco los valores de la factura a autorizar:

        If xEgramar Then
		   punto_vta = CInt(pComprobante.puntoventa.codigo)
		   'punto_vta = string((4-len(punto_vta)),"0")+punto_vta
        Else
			If xCabal or xCilo Then
				punto_vta = cInt (pcomprobante.NUMERADOR.NUMERADOR.CARACTERESPREFIJO)
			
			ELSE
				punto_vta = CInt(pComprobante.BOExtension.PuntoVenta.Codigo)
			End if
        End If
				
		esdelgada		= ObtenerTrDelada(pComprobante)
		
		tipo_cbte 		= ObtenerTipoComprobante(pComprobante)
		
		cbte_nro_calipso = "" : cbte_nro_afip = ""
		
		cbte_nro_calipso = CLng(ObtenerUltimoNumeroCalipso (pComprobante,tipo_cbte,punto_vta))
	
		cbte_nro_afip 	= WSFACT.CompUltimoAutorizado(tipo_cbte, punto_vta)
		
		ControlarExcepcion WSFACT
		For Each v In WSFACT.errores
			senddebug "Factura Electrónica SA - " & v
		Next

		senddebug "Factura Electrónica SA - " & WSFACT.errmsg : senddebug "Factura Electrónica SA - " & WSFACT.errcode
		
		If cbte_nro_afip = ""  or cbte_nro_afip = "0" Then
			cbte_nro_afip = 0                ' no hay comprobantes emitidos
		Else
			if isnull(cbte_nro_afip) Then cbte_nro_afip = 0
			cbte_nro_afip = CLng(cbte_nro_afip)  ' convertir a entero largo								  
		End If
		
	
		ExisteSolicitud = False
		nro_calipso_Actual = 0
		
		'Agregado Ale el 25-03-2020 para recuperar comprobantes que esten en nuestras tablas
		CAE = ""
		
		call ObtenerUltimaSolicitud(pComprobante)  'Obtiene id_Solicitud,nro_calipso_Actual,ExisteSolicitud
		If CAE <> "" Then
				If xWSMTXCA Then
				   cae2 = WSFACT.consultarComprobante(tipo_cbte, punto_vta, nro_calipso_Actual)
				Else
				   cae2 = WSFACT.CompConsultar(tipo_cbte, punto_vta, nro_calipso_Actual)
				End If
				
				ControlarExcepcion WSFACT
				
				ok = WSFACT.AnalizarXml("XmlResponse")
				
				x = ActualizaCamposTransaccion(pcomprobante,WSFACT,punto_vta,tipo_cbte)
							
				x = ActualizaComprobante (pComprobante,WSFACT.Vencimiento,tipo_cbte,punto_vta,nro_calipso_Actual,WSFACT.FechaCbte)
				ReturnValue = True				   
				senddebug "Factura Electrónica SA - Fecha Comprobante:"& WSFACT.FechaCbte : senddebug "Fecha Vencimiento CAE"& WSFACT.Vencimiento :senddebug "Importe Total:"& WSFACT.ImpTotal :	senddebug "Resultado:"& WSFACT.Resultado			
				registraractividadplaunch( "Procesando en Afip" )
				
		Else
		'Fin Agregado Ale el 25-03-2020 para recuperar comprobantes que esten en nuestras tablas

		'50 Si el numero de Calipso y el numero de Afip coinciden entonces comienzo a a procesar el nuevo comprobante	
			If cbte_nro_calipso = cbte_nro_afip Then
				
				If not ExisteSolicitud Then
				
					x = Inserta_Wsfe_Comprobante (pComprobante,tipo_cbte,punto_vta)
					amensaje = ""
					
					x = Inserta_wsfe_Solicitud (pComprobante,nContribuyente,aComputername,ausuario,True,amensaje)
					
					call ObtenerUltimaSolicitud(pComprobante)  'Obtiene id_Solicitud

				End if
				' Creo la Transaccion pidiendo nuevo numero. En dos casos 1 ) Si no existe la Solicitud... en el if anterior ya la creo y 2) si existe la solicitud y en wsfecomprobante no existe numero
				If (ExisteSolicitud and cdbl(nro_calipso_Actual) = 0 ) or  not ExisteSolicitud  Then  		
					
					x =  ObtenerDatosComprobante ( WSFACT, pComprobante,cbte_nro_calipso,esdelgada,tipo_cbte,punto_vta)

				' Habilito reprocesamiento automático (predeterminado):
					WSFACT.Reprocesar = True
	stop
				' Solicito CAE:
					CAE = WSFACT.CAESolicitar()
					ControlarExcepcion WSFACT
					' Imprimo pedido y respuesta XML para depuración (errores de formato)		
					senddebug "Factura Electrónica SA - Resultado"& WSFACT.Resultado : senddebug "CAE"& WSFACT.CAE :	senddebug "Numero de comprobante:"& WSFACT.CbteNro
					senddebug WSFACT.XmlRequest : senddebug WSFACT.XmlResponse :	senddebug "Reproceso:"& WSFACT.Reproceso 
				' Muestro los errores
					amensaje = ""
					If WSFACT.errmsg <> "" Then amensaje = replace(WSFACT.errmsg,chr(10),"")

					xEvento = ""
					If xWSMTXCA Then
					   ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
					   For Each evento In WSFACT.Errores
						  Call CustomMsgBox(evento, vbInformation, "Factura Electrónica")
						  Senddebug "Factura Electrónica SA - " & evento
						  xEvento = xEvento & " - " & evento
					   Next
					Else
					   ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
					   For Each evento In WSFACT.eventos:
						  Call CustomMsgBox(evento, vbInformation, "Factura Electrónica")
						  Senddebug "Factura Electrónica SA " & evento
						  xEvento = xEvento & " - " & evento
					   Next
					End If
					
					If WSFACT.Resultado = "A" and CAE <> "" Then 
						
						x = ActualizaSolicitud (pComprobante,id_Solicitud,amensaje,WSFACT.Resultado)
						
						x = ActualizaCamposTransaccion(pcomprobante,WSFACT,punto_vta,tipo_cbte)
		
						x = ActualizaComprobante (pComprobante,WSFACT.Vencimiento,tipo_cbte,punto_vta,WSFACT.CbteNro,WSFACT.FechaCbte)
						
						senddebug ("Factura Electrónica SA - Numero: " & CStr(punto_vta ) & Right("00000000" & CStr(cbte_nro), 8)):senddebug ("Cae   : "& WSFACT.CAE)
						EVento2 = "Resultado:" & WSFACT.Resultado & " CAE: " & CAE & " Venc: " & WSFACT.Vencimiento & " Obs: " & WSFACT.obs & " Reproceso: " & WSFACT.Reproceso
						registraractividadplaunch( "Procesando en Afip" )
						
						ReturnValue = True

					Else
					   If amensaje = "" Then 
						  amensaje = replace(WSFACT.obs,"'","")
					   Else
						  amensaje = Mid(replace(amensaje,"'",""),1,999)
					   End If
					   
					   x = ActualizaSolicitud (pComprobante,id_Solicitud,amensaje,WSFACT.Resultado)
					   
					   xTime = Replace(Replace(Replace(Replace(CStr(time),":",""),".","")," ",""),";","")
					   
					   If xEvento <> "" Then
						  xtext = xEvento
						  xNombreArchivoXML = nContribuyente & CStr(punto_vta ) & Right("00000000" & CStr(cbte_nro), 8) &  strFechaCae & id_Solicitud&" Eventos " & xTime & ".txt"
						  xArchivoXML = xPathXml & xNombreArchivoXML
						  call EscribirTXT( xArchivoXML, xtext, 0 )
					   End If 

					   xtext = WSFACT.XmlRequest
					   xNombreArchivoXML = nContribuyente & CStr(punto_vta ) & Right("00000000" & CStr(cbte_nro), 8) &  strFechaCae & id_Solicitud&" XmlRequest " & xTime & ".xml"
					   xArchivoXML = xPathXml & xNombreArchivoXML
					   call EscribirTXT( xArchivoXML, xtext, 0 ) 

					   xtext = WSFACT.XmlResponse
					   xNombreArchivoXML = nContribuyente & CStr(punto_vta ) & Right("00000000" & CStr(cbte_nro), 8) &  strFechaCae & id_Solicitud&" XmlResponse " & xTime & ".xml"
					   xArchivoXML = xPathXml & xNombreArchivoXML
					   call EscribirTXT( xArchivoXML, xtext, 0 ) 

					   EVento2 = "Resultado:" & WSFACT.Resultado & " CAE: " & CAE & " Venc: " & WSFACT.Vencimiento & " Obs: " & WSFACT.obs & " Reproceso: " & WSFACT.Reproceso
					   ReturnValue = False
					   Call CustomMsgBox(EVento2, vbInformation + vbOKOnly, "Factura Electrónica")
					   Senddebug "Factura Electrónica SA " & EVento2
					End If
				Else
					'Existe solicitud pero el wsfe Comprobante es mayor a 0. Alguna cosa rara paso
					amensaje = "El Ultimo numero en Afip no concuerda con el de Calipso."
					x = Inserta_wsfe_Solicitud (pComprobante,nContribuyente,aComputername,ausuario,False,amensaje)
					ReturnValue = False
				
				End if
			
			
			End if
			
			
			If cbte_nro_calipso < cbte_nro_afip Then
			   If (ExisteSolicitud and cdbl(nro_calipso_Actual) = 0 ) Then
					cbte_nro = cbte_nro_calipso + 1
					
					If xWSMTXCA Then
					   cae2 = WSFACT.consultarComprobante(tipo_cbte, punto_vta, cbte_nro)
					Else
					   cae2 = WSFACT.CompConsultar(tipo_cbte, punto_vta, cbte_nro)
					End If
					
					ControlarExcepcion WSFACT
					
					ok = WSFACT.AnalizarXml("XmlResponse")
					
					x = ActualizaCamposTransaccion(pcomprobante,WSFACT,punto_vta,tipo_cbte)
								
					x = ActualizaComprobante (pComprobante,WSFACT.Vencimiento,tipo_cbte,punto_vta,cbte_nro,WSFACT.FechaCbte)
					ReturnValue = True				   
					senddebug "Factura Electrónica SA - Fecha Comprobante:"& WSFACT.FechaCbte : senddebug "Fecha Vencimiento CAE"& WSFACT.Vencimiento :senddebug "Importe Total:"& WSFACT.ImpTotal :	senddebug "Resultado:"& WSFACT.Resultado			
					registraractividadplaunch( "Procesando en Afip" )
				Else
					amensaje = "El Ultimo numero en Afip no concuerda con el de Calipso. Debera Recuperar el Comprobante en Forma Manual"
					x = Inserta_wsfe_Solicitud (pComprobante,nContribuyente,aComputername,ausuario,False,amensaje)
					ReturnValue = False
				
				   
				End if
			End if
			
			If cbte_nro_calipso > cbte_nro_afip Then
				amensaje = "El Ultimo numero en Afip no concuerda con el de Calipso."
				x = Inserta_wsfe_Solicitud (pComprobante,nContribuyente,aComputername,ausuario,False,amensaje)
				ReturnValue = False
			End if ' si el Nro de Calipso es igual al de la Afip
		
		End If ' si ya encontre una solicitud con el Numero de Cae en nuestras tablas
	
	else
		amensaje = "Error con Ticket de Acceso a Afip."
		x = Inserta_wsfe_Solicitud (pComprobante,nContribuyente,aComputername,ausuario,False,amensaje)
		ReturnValue = False
	End if
 
 
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'FUNCIONES GENERALES - COMUNES A TODOS LOS CLIENTES
'-------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' @funcion CustomMsgBox
' @descripcion Muestra un mensaje.
'-------------------------------------------------------------------------------
Private Function CustomMsgBox(pMensaje, pBotones, pTitulo)

	If (gInteractivo) Then
	
	    CustomMsgBox = MsgBox(pMensaje, pBotones, pTitulo)
	
	Else

		CustomMsgBox = vbYes
	    Senddebug pMensaje

	End If

End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function ObtenerTicketAccesso (pComprobante,nContribuyente,sServicio,aComputername,ausuario,Certificado,ClavePrivada)
		ObtenerTicketAccesso = False
	
' 10 Verifico en tabla de Ticket de Acceso si tenemos un ticket valido
		aNuevoTicketAccesso = True
		aQuery = "select top 1 * from WSAA_TICKETACCESO where CONTRIBUYENTE = '"&nContribuyente&"' and servicio  = '"&sServicio&"' and Vencimiento > GETDATE() order by ID desc"
		Set xRst = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aQuery )
		Do While Not xRst.EOF
			 aNuevoTicketAccesso = False 
			 aToken = xRst("TOKEN").Value
			 aSign  = xRst("SIGN").Value
			 ObtenerTicketAccesso = True
			 xRst.MoveNext
		Loop
  '20  Generar un Ticket de Requerimiento de Acceso (TRA) para WSFACT porque no encontre en tabla
		If aNuevoTicketAccesso and aToken = "" and aSign = "" Then
			Set WSAA = CreateObject("WSAA")
			' deshabilito errores no manejados (version 2.04 o superior)
			WSAA.LanzarExcepciones = False
			aNuevoTicketAccesso = True
			ttl = 36000 ' tiempo de vida = 10hs hasta expiración
			
			If xWSMTXCA Then
			   tra = WSAA.CreateTRA(sServicio)
			Else
			   tra = WSAA.CreateTRA(sServicio, ttl)
            End If

			ControlarExcepcion WSAA
			senddebug tra
			Path = WSAA.InstallDir + "\conf\"    ' Especificar la ubicacion de los archivos certificado y clave privada
			cacert = WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante

			' Generar el mensaje firmado (CMS)
		    If xUtilizaPathInstalacionSA Then
			   cms  = WSAA.SignTRA(tra, Path + Certificado, Path + ClavePrivada)
            Else		
			   cms = WSAA.SignTRA(tra, Certificado, ClavePrivada)
			End If

			ControlarExcepcion WSAA
			senddebug cms
			' Conectarse con el webservice de autenticación:
			cache 	= ""
			proxy   = ""  
			If xValid Then proxy = "172.16.1.26:8080"
			wrapper = "" ' libreria http (httplib2, urllib2, pycurl)

			cacert 	= WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante

			wsdl  	= "https://wsaa.afip.gov.ar/ws/services/LoginCms?wsdl"

			ObtenerTicketAccesso = WSAA.Conectar(cache, wsdl, proxy, wrapper, cacert) ' Homologación
			ControlarExcepcion WSAA

			If xWSMTXCA Then
			   ' Llamar al web service para autenticar:
               'ta = WSAA.CallWSAA(cms, "https://wsaahomo.afip.gov.ar/ws/services/LoginCms") ' Homologación (cambiar para producción)
               ta = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/LoginCms?wsdl")
               senddebug ta
            Else
			   ' Llamar al web service para autenticar:
			   ta = WSAA.LoginCMS(cms)
			   ControlarExcepcion WSAA
			End If
			
			'x = InsertSQL( "WSAA_TICKETACCESO", "('"&nContribuyente&"','"&sServicio&"',GETDATE(),'"&WSAA.Sign&"','"&WSAA.Token&"',0000000000,DATEADD(N, 600, GETDATE()),'"&aComputername&"','"&ausuario&"','4.0.0.0',9999)", self.WorkSpace )
			aQueryInsert = "INSERT INTO [WSAA_TICKETACCESO] ([CONTRIBUYENTE], [SERVICIO], [GENERACION], [SIGN], [TOKEN], [UNIQUEID], [VENCIMIENTO], [EQUIPO], [USUARIO], [VERSION], [PROCESO]) VALUES ('"&nContribuyente&"','"&sServicio&"',GETDATE(),'"&WSAA.Sign&"','"&WSAA.Token&"',0000000000,DATEADD(N, 600, GETDATE()),'"&aComputername&"','"&ausuario&"','4.0.0.0',9999)"
			Set xRst = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aQueryInsert )
			aToken = WSAA.Token
			aSign	= WSAA.Sign

			cacert = WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante (solo pycurl)

		End if    
	
	End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


' @funcion CreaWSFACT - reemplaza a CreaWSFEv1 o CreaWSMTXCA y ahora depende del parametro de la empresa xWSMTXCA 
' @descripcion Conecta con WSFEv1 o WSMTXCA de sistemas Agiles
'-------------------------------------------------------------------------------	
Public Function CreaWSFACT(aToken,aSign,nContribuyente)
	
		proxy = "" 
		If xValid Then proxy = "172.16.1.26:8080"
        If xWSMTXCA Then
           wsdl = "https://serviciosjava.afip.gob.ar/wsmtxca/services/MTXCAService?wsdl"
        Else
		   wsdl = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL"
        End If
		
		cache = "" 'Path
		wrapper = "" ' libreria http (httplib2, urllib2, pycurl)
		
        If xWSMTXCA Then
		   Set CreaWSFACT = CreateObject("WSMTXCA")
        Else
		   Set CreaWSFACT = CreateObject("WSFEv1")
        End If
		
		senddebug "Factura Electrónica SA - Versión " & CreaWSFACT.Version
		CreaWSFACT.Token 	= aToken
		CreaWSFACT.Sign  	= aSign

		CreaWSFACT.Cuit = nContribuyente     'Cuit del Emisior de la Factura (Cabal, Valid, Sintecrom, ZSA, etc.)
		' deshabilito errores no manejados
		CreaWSFACT.LanzarExcepciones = False
		' Conectar al Servicio Web de Facturación
   
        If xWSMTXCA Then
           ok = CreaWSFACT.Conectar("", wsdl, proxy, "")   ' producción
		   CreaWSFACT.Dummy
		   senddebug "appserver status"& CreaWSFACT.AppServerStatus : senddebug "dbserver status"& CreaWSFACT.DbServerStatus : senddebug "authserver status"& CreaWSFACT.AuthServerStatus
        Else
		   ok = CreaWSFACT.Conectar(cache, wsdl, proxy, wrapper, cacert) ' homologación
		   ControlarExcepcion CreaWSFACT
		   ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
		   CreaWSFACT.Dummy
		   ControlarExcepcion CreaWSFACT
		   senddebug "appserver status"& CreaWSFACT.AppServerStatus : senddebug "dbserver status"& CreaWSFACT.DbServerStatus : senddebug "authserver status"& CreaWSFACT.AuthServerStatus
        End If
End Function		
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	
Sub ControlarExcepcion(Obj)
    ' Nueva funcion para verificar que no haya habido errores:
    On Error GoTo 0

    If obj.Excepcion <> "" and obj.Excepcion <> "OK!" Then
        ' Depuración (grabar a un archivo los detalles del error)
  	   Call CustomMsgBox(obj.Excepcion, vbExclamation, "Factura Electrónica - Excepción")
       Senddebug "Factura Electrónica - Excepción " & obj.Excepcion
    End If

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function	Inserta_Wsfe_Comprobante (pComprobante,tipo_cbte,punto_vta)
				
	aqueryInsertComprobante = "INSERT INTO [WSFE_COMPROBANTE] ([ID], [TIPO], [PUNTOVENTA], [NUMERO], [FECHA], [CAE], [VENCIMIENTOCAE], [NETOGRAVADO], [NETONOGRAVADO], [EXENTO], [IVA], [TRIBUTOS], [TOTAL]) VALUES ('"&pcomprobante.id&"',"& tipo_cbte&","&punto_vta&",0,'','','',0,0,0,0,0,0)"
	Set xRst = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aqueryInsertComprobante )
	Inserta_Wsfe_Comprobante = True
End Function				

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function	Inserta_wsfe_Solicitud (pComprobante,nContribuyente,aComputername,ausuario,Pendiente,aMensaje)
	aPendiente = 0
	aError	   = 0
	If Pendiente Then aPendiente = 1
	If amensaje <> "" Then aError = 1
	aqueryInsertSolicitud = "INSERT INTO [WSFE_SOLICITUD] ([CONTRIBUYENTE], [COMPROBANTE_ID], [PENDIENTE],[ERROR],[MENSAJE],[MOMENTO], [EQUIPO], [USUARIO], [PROCESO], [VERSION]) VALUES ('"&nContribuyente&"','"&pcomprobante.id&"',"&aPendiente&","&aError&",'"&amensaje&"',GETDATE(),'"&aComputername&"','"&ausuario&"',9999,'4000')"
	
	Set xRst = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aqueryInsertSolicitud )
	Inserta_wsfe_Solicitud = True
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function	ObtenerUltimaSolicitud(pComprobante)
		
		aQueryBuscarUltimaSolicitud = "SELECT TOP 1 [A].ID,ISNULL([B].NUMERO,0) AS NUMERO, [B].CAE,[B].VENCIMIENTOCAE FROM [WSFE_SOLICITUD] [A] JOIN [WSFE_COMPROBANTE] [B] ON ([B].[ID] = [A].[COMPROBANTE_ID]) WHERE ([A].[COMPROBANTE_ID] = '"&pcomprobante.id&"') ORDER BY [A].[ID] DESC"
		Set xRstSol = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aQueryBuscarUltimaSolicitud )
		Do While Not xRstSol.EOF
			ExisteSolicitud 	= True
			id_Solicitud     	= xRstSol("ID").Value
			nro_calipso_Actual 	= xRstSol("NUMERO").Value
			CAE					= xRstSol("CAE").Value
			
			xRstSol.MoveNext
		Loop	
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ObtenerUltimoNumeroCalipso (pComprobante,tipo_cbte,punto_vta)
		ObtenerUltimoNumeroCalipso = 0
		aQueryUltimoNro = "select top 1 NUMERO from WSFE_COMPROBANTE where TIPO = "&tipo_cbte&" and PUNTOVENTA = "&punto_vta &" order by NUMERO desc"
		Set xRst = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aQueryUltimoNro )
		Do While Not xRst.EOF
			ObtenerUltimoNumeroCalipso = xRst("NUMERO").Value
			xRst.MoveNext
		Loop
End Function		
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function ActualizaSolicitud (pComprobante,id_Solicitud,amensaje,Resultado)
	ActualizaSolicitud = False
	If Resultado = "A" Then
		aqueryUpdateSolicitud	="UPDATE [WSFE_SOLICITUD] SET [PENDIENTE] = 0, [ERROR] = 0, [MENSAJE] = '' WHERE ([ID] = "&id_Solicitud&")"
	else
		aqueryUpdateSolicitud	="UPDATE [WSFE_SOLICITUD] SET [PENDIENTE] = 0, [ERROR] = 1, [MENSAJE] = '"&amensaje&"' WHERE ([ID] = "&id_Solicitud&")"
	End if
	
	Set xRst 				= RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aqueryUpdateSolicitud )
	ActualizaSolicitud = True

End Function	
	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function	ActualizaComprobante (pComprobante,Vencimiento,tipo_cbte,punto_vta,cbte_nro,fecha_cbte)
		ActualizaComprobante = False
		If xWSMTXCA Then
		   xFecha_Comprobante = Replace(fecha_cbte,"/","")
		   xFecha_Vencimiento = Replace(Vencimiento,"/","")
		Else
		   xFecha_Comprobante = fecha_cbte
		   xFecha_Vencimiento = Vencimiento
        End If
		
		aqueryInsertComprobante = " IF EXISTS(SELECT [ID] FROM [WSFE_COMPROBANTE] WHERE ([ID] = '"&pcomprobante.id&"'))"&_
							   " BEGIN "&_
							   " UPDATE [WSFE_COMPROBANTE] SET [NUMERO] = "&cbte_nro&", [TIPO] = "&tipo_cbte&", [FECHA] = '"&xFecha_Comprobante&"', [CAE] = '"&pComprobante.BOExtension.CAE&"', [VENCIMIENTOCAE] = '"&xFecha_Vencimiento&"', [NETOGRAVADO] = 0, [NETONOGRAVADO] = 0, [EXENTO] = 0, [IVA] = 0, [TRIBUTOS] = 0, [TOTAL] = 0 WHERE ([ID] = '"&pcomprobante.id&"');"&_
							   " END; "&_
							   " ELSE "&_
							   " BEGIN "&_
							   " INSERT INTO [WSFE_COMPROBANTE] ([ID], [TIPO], [PUNTOVENTA], [NUMERO], [FECHA], [CAE], [VENCIMIENTOCAE], [NETOGRAVADO], [NETONOGRAVADO], [EXENTO], [IVA], [TRIBUTOS], [TOTAL]) VALUES ('"&pcomprobante.id&"',"& tipo_cbte&","&punto_vta&","&cbte_nro&",'"&xFecha_Comprobante&"','"&pComprobante.BOExtension.CAE&"','"&xFecha_Vencimiento&"',0,0,0,0,0,0);"&_
							   " END;"
		Set xRst = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aqueryInsertComprobante )
		ActualizaComprobante = True
End Function					 

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' @funcion ActualizaCamposTransaccion
' @Actualiza los campos de Calipso Luego de recibir el CAE
Public Function ActualizaCamposTransaccion(pcomprobante,WSFACT,punto_vta,tipo_cbte)

		ActualizaCamposTransaccion = False			
																													  

'***************************************************************************************************************************************	
'Datos comunes a todas las transacciones y empresas		
		pComprobante.NumeroDocumento = Right("0000" & CStr(punto_vta ),4)  & Right("00000000" & CStr(WSFACT.CbteNro), 8)

'***************************************************************************************************************************************	
'Valid
         If xValid Then
		   pComprobante.BOExtension.CAE	= WSFACT.CAE
		   pComprobante.BOExtension.VencimientoCAE = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.BOExtension.VencimientoCAE) & Right("0" & Month(pComprobante.BOExtension.VencimientoCAE), 2) & Right("0" & Day(pComprobante.BOExtension.VencimientoCAE), 2)
		   pComprobante.BOExtension.CODIGODEBARRAS = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
        End If
					 
'***************************************************************************************************************************************	
'Sintecrom
        If xSintecrom Then
		   pComprobante.BOExtension.CAE	= WSFACT.CAE
		   pComprobante.BOExtension.VENCIMIENTOCAE = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.BOExtension.VENCIMIENTOCAE) & Right("0" & Month(pComprobante.BOExtension.VENCIMIENTOCAE), 2) & Right("0" & Day(pComprobante.BOExtension.VENCIMIENTOCAE), 2)
		   pComprobante.BOExtension.CODIGODEBARRAS = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
        End If

'***************************************************************************************************************************************	
'Cabal
        If xCabal Then
		   pComprobante.BOExtension.CAE	= WSFACT.CAE
		   pComprobante.BOExtension.VencimientoCAE = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.BOExtension.VencimientoCAE) & Right("0" & Month(pComprobante.BOExtension.VencimientoCAE), 2) & Right("0" & Day(pComprobante.BOExtension.VencimientoCAE), 2)
		   pComprobante.BOExtension.CODIGObarracae = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
           pComprobante.NumeroDocumento = CStr(punto_vta )  & Right("00000000" & CStr(WSFACT.CbteNro), 8)
        End If
		
'***************************************************************************************************************************************	
 'ZSA
        If xZSA Then
		   pComprobante.BOExtension.CAE	= WSFACT.CAE
		   pComprobante.BOExtension.VencimientoCAE = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.Boextension.VencimientoCAE) & Right("0" & Month(pComprobante.Boextension.VencimientoCAE), 2) & Right("0" & Day(pComprobante.Boextension.VencimientoCAE), 2)
		   pComprobante.BOExtension.CODIGObarracae = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
        End If
					 
'***************************************************************************************************************************************	
'Egramar
        If xEgramar Then
		   pComprobante.BOExtension.CAE	= WSFACT.CAE
		   pComprobante.BOExtension.fechacae = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.BOExtension.fechacae) & Right("0" & Month(pComprobante.BOExtension.fechacae), 2) & Right("0" & Day(pComprobante.BOExtension.fechacae), 2)
		   pComprobante.BOExtension.codigobarras = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
		   pComprobante.numerodocumento = Right("0000" & CStr(punto_vta ),4) & Right("00000000" & CStr(WSFACT.CbteNro),8)  
		   
		   if pComprobante.BOExtension.CAE <> "" then
	   	      pComprobante.BOExtension.motivocae = "Aprobado"
           End If
		   
        End If

'***************************************************************************************************************************************

'***************************************************************************************************************************************	
'Prodalsa
        If xProdalsa Then
		   pComprobante.BOExtension.CAE	= CStr(WSFACT.CAE)
           pComprobante.BOExtension.VencimientoCAE = cdate(mid(WSFACT.Vencimiento,9,2)&"/"&mid(WSFACT.Vencimiento,6,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.BOExtension.VencimientoCAE) & Right("0" & Month(pComprobante.BOExtension.VencimientoCAE), 2) & Right("0" & Day(pComprobante.BOExtension.VencimientoCAE), 2)
		   stop
		   pComprobante.BOExtension.CodigoBarras = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
		   pComprobante.numerodocumento = Right("0000" & CStr(punto_vta ),4) & Right("00000000" & CStr(WSFACT.CbteNro),8)  
        End If

'***************************************************************************************************************************************

'***************************************************************************************************************************************	
'Binova
        If xBinova Then
		   pComprobante.BOExtension.CAE = CStr(WSFACT.CAE)
		   pComprobante.BOExtension.VencimientoCAE = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.BOExtension.VencimientoCAE) & Right("0" & Month(pComprobante.BOExtension.VencimientoCAE), 2) & Right("0" & Day(pComprobante.BOExtension.VencimientoCAE), 2)
		   pComprobante.BOExtension.CodigoBarra = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
		   pComprobante.NumeroDocumento = Right("0000" & CStr(punto_vta ),4) & Right("00000000" & CStr(WSFACT.CbteNro),8) 
        End If

'***************************************************************************************************************************************

'***************************************************************************************************************************************	
'Stenfar
        If xStenfar Then
		   pComprobante.BOExtension.CAE = CStr(WSFACT.CAE)
		   pComprobante.BOExtension.VencimientoCAE = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   pComprobante.BOExtension.FECHAVTOCAE    = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.BOExtension.VencimientoCAE) & Right("0" & Month(pComprobante.BOExtension.VencimientoCAE), 2) & Right("0" & Day(pComprobante.BOExtension.VencimientoCAE), 2)
		   pComprobante.BOExtension.CODIGObarracae = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
		   pComprobante.NumeroDocumento = Right("0000" & CStr(punto_vta ),4) & Right("00000000" & CStr(WSFACT.CbteNro),8) 
			
			If ObtieneTransicion(pComprobante) <> "" then
   			  	   registraractividadplaunch("Ejecutando Transicion")
	 	   	  	   x =  EjecutarTransicion( pComprobante, ObtieneTransicion(pComprobante) )
	 	   	  	   If x Then
  					  ActualizaCamposTransaccion = True
  			  		  registraractividadplaunch("Transicion Ejecutada")
		 		   Else
		 			  ActualizaCamposTransaccion = False
  			  		  registraractividadplaunch("Transicion Ejecutada")
		 		   End if
			End if		 
		
		End If

'***************************************************************************************************************************************
		'***************************************************************************************************************************************	
'Cilo
        If xCilo Then
		   pComprobante.BOExtension.CAE = CStr(WSFACT.CAE)
		   pComprobante.BOExtension.VencimientoCAE = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   pComprobante.BOExtension.FECHAVTOCAE = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.BOExtension.VencimientoCAE) & Right("0" & Month(pComprobante.BOExtension.VencimientoCAE), 2) & Right("0" & Day(pComprobante.BOExtension.VencimientoCAE), 2)
		   pComprobante.BOExtension.CODIGOBARRACAE = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
		   pComprobante.NumeroDocumento = Right("0000" & CStr(punto_vta ),4) & Right("00000000" & CStr(WSFACT.CbteNro),8) 
        End If

'***************************************************************************************************************************************



		ActualizaCamposTransaccion = True

				
End Function

					 
' @funcion ObtenerTipoComprobante
' @descripcion Obtiene el código de tipo de comprobante.
'-------------------------------------------------------------------------------
Private Function ObtenerTipoComprobante(ByRef pComprobante)

	Dim iTipoComprobante
	
	' Establece el tipo de comprobante en base a la letra de la factura.
	If (pComprobante.Nota = "A") Then
	
	    iTipoComprobante = 1
	
	ElseIf (pComprobante.Nota = "B") Then
	
		iTipoComprobante = 6
	
	ElseIf (pComprobante.Nota = "C") Then
	
		iTipoComprobante = 11
	
	End If
		' Establece el tipo de comprobante de acuerdo al tipo de factura.
	sClassName = ClassName(pComprobante)
	If (sClassName = "TRFACTURAVENTA") Or (sClassName = "TDFACTURAVENTA") Then
	
	    iTipoComprobante = iTipoComprobante + 0

	ElseIf (sClassName = "TRDEBITOVENTA") Or (sClassName = "TDDEBITOVENTA") Then 
	
	    iTipoComprobante = iTipoComprobante + 1
	
	ElseIf (sClassName = "TRCREDITOVENTA") Or (sClassName = "TDCREDITOVENTA") Then
	
	    iTipoComprobante = iTipoComprobante + 2
	
	End If

'------Activado el 1/7/2019 para MiPyMEs o Factura de Crédito
	If pComprobante.BOExtension.MiPyMEs Then
	   iTipoComprobante = iTipoComprobante + 200
	End If
'------Fin Activado el 1/7/2019 para MiPyMEs o Factura de Crédito
	
	' Establece el resultado.
	ObtenerTipoComprobante = iTipoComprobante
	
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' @funcion ObtenerVencimientoPago
' @descripcion Obtiene la fecha de vencimiento del pago.
'-------------------------------------------------------------------------------
Private Function ObtenerVencimientoPago(ByRef pComprobante)

	Dim dVencimiento
	' Inicializa el vencimiento.
	dVencimiento = pComprobante.FechaActual

	' Busca el vencimiento en el tipo de pago.
	If (Not pComprobante.TipoPago Is Nothing) Then

	    For Each oItemTipoPago In pComprobante.TipoPago.ItemsTipoPago

           If xZSA or xValid Then
			  dVencimiento = dVencimiento + oItemTipoPago.DiasVencimiento 
           Else
			  dVencimiento = dVencimiento + oItemTipoPago.Periodo 
           End If
		   
			Exit For
		Next
	Else
	
	    dVencimiento = dVencimiento + 10
	End If
	ObtenerVencimientoPago = year(dVencimiento) & right("00" & month(dVencimiento),2) & right("00" & day(dVencimiento),2)
	' Devuelve el vencimiento.
'	ObtenerVencimientoPago = dVencimiento

End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' @funcion ArmoCodigoBarra
' @descripcion Armo el código de barras.
'-------------------------------------------------------------------------------
Private Function ArmoCodigoBarra(pComprobante,strcodigocomprobante,strcuit,strpuntoventa,strcae,strFecha)
	If xValid or xStenfar or xSintecrom Then
		xSQL = "SELECT calipso.FN_CODIGOBARRA_CAE('" & strcuit & strcodigocomprobante & strPuntoVenta & strCAE & strFecha & "') AS CODIGO "
	Else
		xSQL = "SELECT dbo.FN_CODIGOBARRA_CAE('" & strcuit & strcodigocomprobante & strPuntoVenta & strCAE & strFecha & "') AS CODIGO "
    End if  
	Set xRst = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), xSQL )
    Do While Not xRst.EOF
         ArmoCodigoBarra = xRst("CODIGO").Value
         xRst.MoveNext
    Loop

End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' @funcion esdelgada
' @descripcion Determina si una transaccion es delgada o no
'-------------------------------------------------------------------------------

Public Function ObtenerTrDelada (pComprobante)

	sClassName = ClassName(pComprobante)
	If (sClassName = "TDCREDITOVENTA") Or (sClassName = "TDFACTURAVENTA") Or (sClassName = "TDDEBITOVENTA") or (sClassName = "TDFACTURAANTICIPO") Then
	    ObtenerTrDelada = True
	else
		ObtenerTrDelada = False
	End If

End Function




'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'FUNCIONES ESPECIFICAS PARA CADA CLIENTE POR LOS ID DE CADA UNO
'-------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' @funcion ObtenerObjetoIva
' @descripcion Obtengo el Objeto iva tanto de tr delgada como la comun.
Private Function ObtenerObjetoIva(oItemComprobante,esdelgada)
		set ObtenerObjetoIva = Nothing
		If esdelgada Then
			set xview = NewCompoundView( oItemComprobante, "TDIMPUESTOTRANSACCION", oItemComprobante.Workspace, nil, True )
			xview.addfilter(NewFilterSpec( xview.columnfrompath("BOOWNER"), "=", oItemComprobante ))
			For each item in xview.VIEWITEMS
					if item.bo.definicionimpuesto.impuesto.codigo = "010" then
								set oIva = item.bo
					end if										 
			Next
		else
			Set oIva = GetImpuestoTrPorCodigo(oItemComprobante, "010", "I").Owner
		End if
		set ObtenerObjetoIva = oIva
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' @funcion ObtenerTotalItem
' @descripcion Obtengo el Total de los Items delgada o no
'-------------------------------------------------------------------------------				
Private Function ObtenerTotalItem (oItemComprobante,esdelgada)
	ObtenerTotalItem = 0
	If esdelgada Then
		ObtenerTotalItem = oItemComprobante.Total_Importe
	Else
		ObtenerTotalItem = oItemComprobante.Total.Importe
	End if
End Function		
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' @funcion ObtenerIvaItem
' @descripcion Obtengo el Iva Total de los Items delgada o no
'-------------------------------------------------------------------------------				
Private Function ObtenerIvaItem (oIva,esdelgada)
	ObtenerIvaItem = 0
	If esdelgada Then
		ObtenerIvaItem = oIva.Importe
	Else
		ObtenerIvaItem = oIva.Valor.Importe
	End if
End Function		
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' @funcion ObtenerTotal_Otros_Tributos
' @descripcion Obtengo el Total de Otros Tributos delgada o no
'-------------------------------------------------------------------------------
Private Function ObtenerTotal_Otros_Tributos (pComprobante,esdelgada)
	ObtenerTotal_Otros_Tributos = 0
	Total_Otros_Tributos		= 0
	If esdelgada Then
			For Each oImpuesto In pComprobante.Impuestos
			   If (oImpuesto.Importe > 0) Then
				  If oImpuesto.DefinicionImpuesto.Impuesto.Codigo <> "010" Then	
					 Total_Otros_Tributos = Total_Otros_Tributos + oImpuesto.Importe
				  End If
			   End If
			Next
	else
			For Each oImpuesto In pComprobante.ImpuestosTransaccion
			   If (oImpuesto.valor.Importe > 0) Then
				  If oImpuesto.DefinicionImpuesto.Impuesto.Codigo <> "010" Then	
					 Total_Otros_Tributos = Total_Otros_Tributos + oImpuesto.valor.Importe
				  End If
			   End If
			Next
	End if
	ObtenerTotal_Otros_Tributos = Total_Otros_Tributos
End Function	

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' @funcion ObtenerDatosComprobante
' @descripcion Obtengo el Total de Otros Tributos delgada o no
'-------------------------------------------------------------------------------
Private Function ObtenerDatosComprobante ( WSFACT, pComprobante,cbte_nro_calipso,esdelgada,tipo_cbte,punto_vta)
				ObtenerDatosComprobante = False
STOP				
                If xWSMTXCA Then
				   fecha = year(pComprobante.FechaActual) &"-"& right("00" & month(pComprobante.FechaActual),2) &"-"& right("00" & day(pComprobante.FechaActual),2)
                Else
				   fecha = year(pComprobante.FechaActual) & right("00" & month(pComprobante.FechaActual),2) & right("00" & day(pComprobante.FechaActual),2)
                End If
				 
				concepto = 1  'Bienes '1 - Productos / 2 - Servicios / 3 - Productos y servicios

				tipo_doc = 80

'*******************************************************************************
'*******************************************************************************
                '------ Modificado el 22/12/2019 para Prodalsa porque usa un solo destinatario para consumidor final
				'nro_doc = pComprobante.Destinatario.EnteAsociado.Cuit
				If xProdalsa Then
                   If pComprobante.Destinatario.Codigo = "9999" Then 'CONSUMIDOR FINAL
	                  nro_doc = pComprobante.BOExtension.EMPLEADO.EnteAsociado.CUIT
	               Else
	                  nro_doc= pComprobante.Destinatario.EnteAsociado.CUIT
	               End If
				Else
				   nro_doc = pComprobante.Destinatario.EnteAsociado.Cuit
                End If
				If xStenfar Then
					
					if pComprobante.nombreDestinatario = "EMPLEADOS" and pComprobante.Total.Importe < 1000 Then
	
						tipo_doc = 99
						nro_doc	 = 0
					else
						tipo_doc    = 80
						nro_doc 	= pComprobante.Destinatario.EnteAsociado.Cuit
					End if   
					' Es Mercado Libre y Consumidor Final
	
					If pComprobante.boextension.POSICIONIVACLIENTE.name = "CONSUMIDOR FINAL" Then
						If pComprobante.Total.Importe < 10000 Then
							tipo_doc = 99
							nro_doc	 = 0
						Else
							If pComprobante.Tipotransaccion.OTP = "TRFACTURAVENTA" Then
							 If  pComprobante.boextension.ML_NOMBRE <> "" Then
								 If  ucase(pComprobante.boextension.ML_TIPODOC) = "CUIT" Then tipo_doc = 80
								 If  ucase(pComprobante.boextension.ML_TIPODOC) = "DNI" Then tipo_doc = 96
								 nro_doc 	= int(pComprobante.boextension.ML_NUMERODOC)
							  End if
							End if
							If  pComprobante.Tipotransaccion.OTP = "TRCREDITOVENTA" Then
							 If  pComprobante.boextension.ML_NOMBRE <> "" Then
								 If  ucase(pComprobante.boextension.ML_TIPODOC) = "CUIT" Then tipo_doc = 80
								 If  ucase(pComprobante.boextension.ML_TIPODOC) = "DNI" Then  tipo_doc = 96
								 nro_doc 	= pComprobante.boextension.ML_NUMERODOC
							  End if
							 End if			
						End if
					End if
				End if
					
	            '------ Fin Modificado el 22/12/2019 para Prodalsa porque usa un solo destinatario para consumidor final
'*******************************************************************************
'*******************************************************************************

				nro_doc = replace(nro_doc,"-","")
				nro_doc = replace(nro_doc,".","")
				nro_doc = replace(nro_doc," ","")

				cbte_nro = cbte_nro_calipso + 1

				cbt_desde = cbte_nro: cbt_hasta = cbte_nro

				imp_iva = "0": imp_trib = "0": imp_op_ex = "0.00"

				fecha_cbte = fecha
				If tipo_cbte = 201 Then   ' Solo FCE
				   fecha_venc_pago = ObtenerVencimientoPago(pComprobante)
                   If xWSMTXCA Then
				      fecha_venc_pago = MID(fecha_venc_pago,1,4)&"-"&MID(fecha_venc_pago,5,2)&"-"&MID(fecha_venc_pago,7,2)
				   End if
				else
				   fecha_venc_pago = ""
				End if
				' Fechas del período del servicio facturado (solo si concepto = 1?)
				 fecha_serv_desde = "": fecha_serv_hasta = ""
				If Concepto = 80 Then  fecha_serv_desde = fecha: fecha_serv_hasta = fecha
					   
				If Not FuerzaDolarizacionComprobante Then
				   moneda_id = ObtenerMoneda(pComprobante)
				End If

                 moneda_ctz = pComprobante.Cotizacion
				If FuerzaDolarizacionComprobante Then
				   moneda_ctz = CotizacionDolar
				End If
				
				imp_op_ex = 0 : imp_tot_conc = 0 : TotalNetoNoGravado = 0 :	TotalNetoGravado = 0
				base_imp_0 = 0 : valoriva_0 = 0 : base_imp_105 = 0 : valoriva_105 = 0
				base_imp_21 = 0 : valoriva_21 = 0 :	base_imp_27 = 0 : valoriva_27 = 0

'*******************************************************************************
'*******************************************************************************
                '------ Modificado el 23/04/2021 para que contemple los clientes con factura A pero con posición RI Tierra del Fuego
				' Obtiene la posición de IVA si es cliente.	
			    If CLASSNAME(pComprobante.Destinatario) = "CLIENTE" Then
                   Set oPosicionCliente = GetPosicionImpuestoPorCodigo(pComprobante.Destinatario, "010")
                   xCodigoPosicionCliente = oPosicionCliente.PosicionImpuesto.Codigo
                Else
                   xCodigoPosicionCliente = "MO"
                End If
                '------ Fin Modificado el 23/04/2021 para que contemple los clientes con factura A pero con posición RI Tierra del Fuego
'*******************************************************************************
'*******************************************************************************

				' Calcula los netos y el IVA.	
				For Each oItemComprobante In pComprobante.ItemsTransaccion

				' Obtiene la posición de IVA.	
					Set oPosicion = GetPosicionImpuestoPorCodigo(oItemComprobante.Referencia, "010")
					
'*******************************************************************************
'*******************************************************************************
                '------ Modificado el 23/04/2021 para que contemple los clientes con factura A pero con posición RI Tierra del Fuego
					'If ( mid(oPosicion.PosicionImpuesto.Codigo,1,2) = "EX") Then		
					If ( mid(oPosicion.PosicionImpuesto.Codigo,1,2) = "EX") or (xCodigoPosicionCliente = "RT") Then		
                '------ Fin Modificado el 23/04/2021 para que contemple los clientes con factura A pero con posición RI Tierra del Fuego
'*******************************************************************************
'*******************************************************************************
							imp_op_ex = imp_op_ex + ObtenerTotalItem (oItemComprobante,esdelgada)		
					Else
						' Obtiene el IVA y lo clasifica.
						set oIva = ObtenerObjetoIva(oItemComprobante,esdelgada)
						If (oPosicion.Coeficiente = 0) Then
							id_0 = 3
'Activé el 21/08/2020 porque daba errores en los items con IVA 0
							imp_tot_conc = Round(imp_tot_conc +  ObtenerTotalItem (oItemComprobante,esdelgada),4)
							base_imp_0 = Round(base_imp_0 +  ObtenerTotalItem (oItemComprobante,esdelgada),4)
							valoriva_0 = Round(valoriva_0 + ObtenerIvaItem (oIva,esdelgada),4)
					
						ElseIf (oPosicion.Coeficiente = 10.5) Then
							id_105 = 4
							If xValid Then	
								If pComprobante.Nota = "B" Then	
									base_imp_105 = Round(base_imp_105 +  ObtenerTotalItem (oItemComprobante,esdelgada) / (1 + (oPosicion.Coeficiente/100)) ,4)
									valoriva_105 = Round(valoriva_105 + round(ObtenerTotalItem (oItemComprobante,esdelgada) / ((1+(oPosicion.Coeficiente/100))/(oPosicion.Coeficiente/100)),2),2)
									TotalNetoGravado = TotalNetoGravado +  round((ObtenerTotalItem (oItemComprobante,esdelgada) / (1+(oPosicion.Coeficiente/100))),4)
								else
									base_imp_105 = Round(base_imp_105 +  ObtenerTotalItem (oItemComprobante,esdelgada) ,4)
									valoriva_105 = Round(valoriva_105 + ObtenerIvaItem (oIva,esdelgada),4)
									TotalNetoGravado = TotalNetoGravado +  ObtenerTotalItem (oItemComprobante,esdelgada)
								End if
							else
								base_imp_105 = Round(base_imp_105 +  ObtenerTotalItem (oItemComprobante,esdelgada) ,4)
								valoriva_105 = Round(valoriva_105 + ObtenerIvaItem (oIva,esdelgada),4)
								TotalNetoGravado = TotalNetoGravado +  ObtenerTotalItem (oItemComprobante,esdelgada)
							End if
					
						ElseIf (oPosicion.Coeficiente = 21) Then
							id_21 = 5
							If xValid Then	
								If pComprobante.Nota = "B" Then	
									base_imp_21 = Round(base_imp_21 +  ObtenerTotalItem (oItemComprobante,esdelgada) / (1 + (oPosicion.Coeficiente/100)) ,4)
									 valoriva_21 = Round(valoriva_21 + round(ObtenerTotalItem (oItemComprobante,esdelgada) / ((1+(oPosicion.Coeficiente/100))/(oPosicion.Coeficiente/100)),2),2)
									TotalNetoGravado = TotalNetoGravado +  round((ObtenerTotalItem (oItemComprobante,esdelgada) / (1+(oPosicion.Coeficiente/100))),4)
								else
									base_imp_21 = Round(base_imp_21 +  ObtenerTotalItem (oItemComprobante,esdelgada) ,4)
									valoriva_21 = Round(valoriva_21 + ObtenerIvaItem (oIva,esdelgada),4)
									TotalNetoGravado = TotalNetoGravado +  ObtenerTotalItem (oItemComprobante,esdelgada)
								End if
							
							else
								base_imp_21 = Round(base_imp_21 +  ObtenerTotalItem (oItemComprobante,esdelgada) ,4)
								valoriva_21 = Round(valoriva_21 + ObtenerIvaItem (oIva,esdelgada),4)
								TotalNetoGravado = TotalNetoGravado +  ObtenerTotalItem (oItemComprobante,esdelgada)
							End if

						ElseIf (oPosicion.Coeficiente = 27) Then
							id_27 = 6
							If xValid Then	
								If pComprobante.Nota = "B" Then	
									base_imp_27 = Round(base_imp_27 +  ObtenerTotalItem (oItemComprobante,esdelgada) / (1 + oPosicion.Coeficiente) ,4)
									valoriva_27 = Round(valoriva_27 + round(ObtenerTotalItem (oItemComprobante,esdelgada) / ((1+(oPosicion.Coeficiente/100))/(oPosicion.Coeficiente/100)),2),2)
									TotalNetoGravado = TotalNetoGravado +  round((ObtenerTotalItem (oItemComprobante,esdelgada) / (1+oPosicion.Coeficiente)),4)
								else
									base_imp_27 = Round(base_imp_27 +  ObtenerTotalItem (oItemComprobante,esdelgada) ,4)
									valoriva_27 = Round(valoriva_27 + ObtenerIvaItem (oIva,esdelgada),4)
									TotalNetoGravado = TotalNetoGravado +  ObtenerTotalItem (oItemComprobante,esdelgada)
								End if
							else
								base_imp_27 = Round(base_imp_27 +  ObtenerTotalItem (oItemComprobante,esdelgada) ,4)
								valoriva_27 = Round(valoriva_27 + ObtenerIvaItem (oIva,esdelgada),4)
								TotalNetoGravado = TotalNetoGravado +  ObtenerTotalItem (oItemComprobante,esdelgada)
							End if
			
						End If
						
					End If 'Si es Excento o no
				
				Next

				Total_Otros_Tributos = 0 
				Total_Otros_Tributos = ObtenerTotal_Otros_Tributos (pComprobante,esdelgada)
				If FuerzaDolarizacionComprobante Then
				   moneda_ctz = CotizacionDolar
				End If
				
				imp_neto = Round(TotalNetoGravado,2)
				imp_iva  = Round(valoriva_0 + valoriva_105 + valoriva_21 + valoriva_27,2)
				imp_trib = Round(Total_Otros_Tributos,2) 'Impuestos como IIBB y Percepciones ver como hacer
				Imp_TotalCalculado = imp_neto + imp_iva + Imp_Trib + imp_tot_conc + imp_op_ex
				 
'*******************************************************************************
'*******************************************************************************
                '------ Agregado 22/12/2019 para WSMTXCA porque neceita SubTotal
				If esdelgada then
					imp_subtotal = Round(pComprobante.SubTotal_Importe,2)
				else
					imp_subtotal = Round(pComprobante.SubTotal.Importe,2)
				End if
                '------ Fin Agregado 22/12/2019 para WSMTXCA porque neceita SubTotal
'*******************************************************************************
'*******************************************************************************

				If esdelgada then
					imp_total = Round(pComprobante.Total_Importe,2)
				else
					imp_total = Round(pComprobante.Total.Importe,2)
				End if


'*******************************************************************************
'*******************************************************************************
                '------ Agregado 08/03/2021 para Forzar Dolarización de Comprobantes
				If FuerzaDolarizacionComprobante Then
                                   If imp_neto > 0 Then
				      imp_neto = Round(imp_neto/CotizacionDolar,2)
				   End If
                                   If imp_iva > 0 Then
				      imp_iva  = Round(imp_iva/CotizacionDolar,2)
				   End If
                                   If imp_trib > 0 Then
				      imp_trib = Round(imp_trib/CotizacionDolar,2)
				   End If
                                   If Imp_TotalCalculado > 0 Then
				      Imp_TotalCalculado = Round(Imp_TotalCalculado/CotizacionDolar,2)
				   End If
                                   If imp_subtotal > 0 Then
				      imp_subtotal = Round(imp_subtotal/CotizacionDolar,2)
				   End If
                                   If imp_total > 0 Then
				      imp_total = Round(imp_total/CotizacionDolar,2)
				   End If
                                   If imp_tot_conc > 0 Then
                                      imp_tot_conc = Round(imp_tot_conc/CotizacionDolar,2)
				   End If
                                   If imp_op_ex > 0 Then
                                      imp_op_ex = Round(imp_op_ex/CotizacionDolar,2)
				   End If
				End If

                '------ Fin Agregado 08/03/2021 para Forzar Dolarización de Comprobantes
'*******************************************************************************
'*******************************************************************************
				If xValid Then
					If Round(CDbl(round(imp_total,2)) - CDbl(round(Imp_TotalCalculado,2)),2)> 0 Then 
					   imp_op_ex = imp_op_ex + Round(CDbl(round(imp_total,2)) - CDbl(round(Imp_TotalCalculado,2)),2)	
					End if
				End if
        		
				If CDbl(round(imp_total,2)) - CDbl(round(Imp_TotalCalculado,2)) <= imp_Tolerancia Then  Imp_Total = Imp_TotalCalculado
				                       
                If xWSMTXCA Then
				stop
				   obs = "" 'Acá pueden ir las observaciones que se necesiten en la factura 
				   ObtenerDatosComprobante = WSFACT.CrearFactura(concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta,cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto,imp_subtotal, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago,fecha_serv_desde, fecha_serv_hasta, moneda_id, moneda_ctz, obs)
                Else
				   ObtenerDatosComprobante = WSFACT.CrearFactura(concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta,cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto,imp_iva, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago,fecha_serv_desde, fecha_serv_hasta, moneda_id, moneda_ctz)
                End If


				If (tipo_cbte = 203 OR tipo_cbte = 202) and pComprobante.BOExtension.MiPyMEs Then  ' Si es una Nota de Credito y es Mipymes obligo a que pongo comprobante asociado
				   'set xFacturaVinculada = pComprobante.VinculoTr
				   If xZSA or xEgramar or xProdalsa or xCabal Then
						set xFacturaVinculada = pComprobante.boextension.cpvinculado.TrOriginante
				   Else
						set xFacturaVinculada = pComprobante.VinculoTr
				   End if
				   If not xFacturaVinculada is Nothing Then
				   	  tipo_FacVinc    = 201
				      If xEgramar Then
				   	     pto_vta_FacVinc = CInt(xFacturaVinculada.PuntoVenta.Codigo)
                      Else

			              If xCabal Then
				             pto_vta_FacVinc = cInt (xFacturaVinculada.NUMERADOR.NUMERADOR.CARACTERESPREFIJO)
			
			               ELSE
				   	     pto_vta_FacVinc = CInt(xFacturaVinculada.BOExtension.PuntoVenta.Codigo)
			              End if


					  End If
				   	  nro_FacVinc     = cint(right(xFacturaVinculada.numerodocumento,8))
					  fecha_Factura   = year(xFacturaVinculada.FechaActual) & right("00" & month(xFacturaVinculada.FechaActual),2) & right("00" & day(xFacturaVinculada.FechaActual),2)
 
                      If xWSMTXCA Then
						 ObtenerDatosComprobante   = WSFACT.AgregarCmpAsoc(tipo_FacVinc, pto_vta_FacVinc, nro_FacVinc)
				   	  Else   
						 ObtenerDatosComprobante   = WSFACT.AgregarCmpAsoc(tipo_FacVinc, pto_vta_FacVinc, nro_FacVinc,nContribuyente,fecha_Factura)
                      End If

				   End if

'*******************************************************************************
'*******************************************************************************
                '------ Agregado 22/12/2019 para WSMTXCA porque neceita vincular comprobantes
'*******************************************************************************
                '------ Adaptado 28/07/2020 porque neceita vincular comprobantes o fechas de periodo para las ND y NC
                '------ Si no hay Comprobante Asociado, se asignan las fechas.
				'ElseIf (tipo_cbte = 3 OR tipo_cbte = 2) and xWSMTXCA and Not pComprobante.BOExtension.MiPyMEs Then
				ElseIf (tipo_cbte = 3 OR tipo_cbte = 2  OR tipo_cbte = 7 OR tipo_cbte = 8) and Not pComprobante.BOExtension.MiPyMEs Then
stop
				   'set xFacturaVinculada = pComprobante.VinculoTr
                   set xComprobanteAsociado = nothing
						If xProdalsa or xZSA or xCabal or xEgramar Then
							  If Not pComprobante.boextension.AFIP_COMPROBANTE_ASOCIADO is Nothing Then
								 set xComprobanteAsociado = pComprobante.boextension.AFIP_COMPROBANTE_ASOCIADO
							  End If			   
						End if
						If xValid Then
							set xComprobanteAsociado = pComprobante.VinculoTr
						End if
						If not xComprobanteAsociado is Nothing Then

							tipo_DocAsoc = ObtenerTipoComprobante(xComprobanteAsociado)
							  If xEgramar Then
									pto_vta_DocAsoc = CInt(xComprobanteAsociado.PuntoVenta.Codigo)
							  Else
									If xCabal Then
												 pto_vta_DocAsoc = CInt(xComprobanteAsociado.NUMERADOR.NUMERADOR.CARACTERESPREFIJO)
									Else
												 pto_vta_DocAsoc = CInt(xComprobanteAsociado.BOExtension.PuntoVenta.Codigo)
									End if
							  End If
							'nro_DocAsoc     = cint(right(xComprobanteAsociado.numerodocumento,8))
							nro_DocAsoc     = clng(right(xComprobanteAsociado.numerodocumento,8))
							fecha_DocAsoc   = year(xComprobanteAsociado.FechaActual) & right("00" & month(xComprobanteAsociado.FechaActual),2) & right("00" & day(xComprobanteAsociado.FechaActual),2)
 
							  If xWSMTXCA Then
								 ObtenerDatosComprobante   = WSFACT.AgregarCmpAsoc(tipo_DocAsoc, pto_vta_DocAsoc, nro_DocAsoc)
							  Else   
								 ObtenerDatosComprobante   = WSFACT.AgregarCmpAsoc(tipo_DocAsoc, pto_vta_DocAsoc, nro_DocAsoc,nContribuyente,fecha_DocAsoc)
							  End If

						Else
							  If xProdalsa or xZSA or xCabal or xEgramar Then

								 fecha_desde_DocAsoc = Cdate(pComprobante.boextension.AFIP_PERIODO_DESDE) 
								 If xWSMTXCA Then
									fecha_desde_DocAsoc = DatePart("yyyy", fecha_desde_DocAsoc) &"-"& string(2-len(DatePart("m", fecha_desde_DocAsoc)),"0") & DatePart("m", fecha_desde_DocAsoc) &"-"& string(2-len(DatePart("d", fecha_desde_DocAsoc)),"0") & DatePart("d", fecha_desde_DocAsoc)	
								 Else
									fecha_desde_DocAsoc = DatePart("yyyy", fecha_desde_DocAsoc) & string(2-len(DatePart("m", fecha_desde_DocAsoc)),"0") & DatePart("m", fecha_desde_DocAsoc) & string(2-len(DatePart("d", fecha_desde_DocAsoc)),"0") & DatePart("d", fecha_desde_DocAsoc)	
								 End If
								 
								 fecha_hasta_DocAsoc = Cdate(pComprobante.boextension.AFIP_PERIODO_HASTA)
								 If xWSMTXCA Then
									fecha_hasta_DocAsoc = DatePart("yyyy", fecha_hasta_DocAsoc) &"-"& string(2-len(DatePart("m", fecha_hasta_DocAsoc)),"0") & DatePart("m", fecha_hasta_DocAsoc) &"-"& string(2-len(DatePart("d", fecha_hasta_DocAsoc)),"0") & DatePart("d", fecha_hasta_DocAsoc)	
								 Else
									fecha_hasta_DocAsoc = DatePart("yyyy", fecha_hasta_DocAsoc) & string(2-len(DatePart("m", fecha_hasta_DocAsoc)),"0") & DatePart("m", fecha_hasta_DocAsoc) & string(2-len(DatePart("d", fecha_hasta_DocAsoc)),"0") & DatePart("d", fecha_hasta_DocAsoc)	
								 End If
								 
								 ObtenerPeriodoComprobante = WSFACT.AgregarPeriodoComprobantesAsociados(fecha_desde_DocAsoc, fecha_hasta_DocAsoc)				   

							  End if
							  If xValid or xStenfar or xSintecrom or xCilo Then
									fecha_desde_DocAsoc 	= pComprobante.FechaActual - 15
									fecha_desde_DocAsoc 	= year(fecha_desde_DocAsoc) & right("00" & month(fecha_desde_DocAsoc),2) & right("00" & day(fecha_desde_DocAsoc),2)
									fecha_hasta_DocAsoc 	= pComprobante.FechaActual
									fecha_hasta_DocAsoc   	= year(fecha_hasta_DocAsoc) & right("00" & month(fecha_hasta_DocAsoc),2) & right("00" & day(fecha_hasta_DocAsoc),2)
									ObtenerPeriodoComprobante = WSFACT.AgregarPeriodoComprobantesAsociados(fecha_desde_DocAsoc, fecha_hasta_DocAsoc)
							  End if
						End if
                '------ Fin Adaptado 28/07/2020 porque neceita vincular comprobantes o fechas de periodo para las ND y NC
'*******************************************************************************
                '------ Fin Agregado 22/12/2019 para WSMTXCA porque neceita vincular comprobantes
'*******************************************************************************
'*******************************************************************************

				End if
				

'*******************************************************************************
'*******************************************************************************
                '------ Agregado 08/03/2021 para Forzar Dolarización de Comprobantes

				If FuerzaDolarizacionComprobante Then
                                   If valoriva_0 > 0 Then 	ObtenerDatosComprobante = WSFACT.AgregarIva(id_0, Round((base_imp_0/CotizacionDolar),2), Round((valoriva_0/CotizacionDolar),2))
				   If valoriva_105 > 0 Then 	ObtenerDatosComprobante = WSFACT.AgregarIva(id_105, Round((base_imp_105/CotizacionDolar),2), Round((valoriva_105/CotizacionDolar),2))
				   If valoriva_21 > 0 Then 	ObtenerDatosComprobante = WSFACT.AgregarIva(id_21, Round((base_imp_21/CotizacionDolar),2), Round((valoriva_21/CotizacionDolar),2))
				   If valoriva_27 > 0 Then 	ObtenerDatosComprobante = WSFACT.AgregarIva(id_27, Round((base_imp_27/CotizacionDolar),2), Round((valoriva_27/CotizacionDolar),2))
                                Else				
                                
                '------ Fin Agregado 08/03/2021 para Forzar Dolarización de Comprobantes
'*******************************************************************************
'*******************************************************************************
                                   If valoriva_0 > 0 Then 	ObtenerDatosComprobante = WSFACT.AgregarIva(id_0, Round(base_imp_0,2), Round(valoriva_0,2))
				   If valoriva_105 > 0 Then 	ObtenerDatosComprobante = WSFACT.AgregarIva(id_105, Round(base_imp_105,2), Round(valoriva_105,2))
				   If valoriva_21 > 0 Then 	ObtenerDatosComprobante = WSFACT.AgregarIva(id_21, Round(base_imp_21,2), Round(valoriva_21,2))
				   If valoriva_27 > 0 Then 	ObtenerDatosComprobante = WSFACT.AgregarIva(id_27, Round(base_imp_27,2), Round(valoriva_27,2))
'*******************************************************************************
'*******************************************************************************
                '------ Agregado 08/03/2021 para Forzar Dolarización de Comprobantes
				End If
                '------ Fin Agregado 08/03/2021 para Forzar Dolarización de Comprobantes
'*******************************************************************************
'*******************************************************************************


				'1	Impuestos nacionales     '2	Impuestos provinciales     '3	Impuestos municipales     '4	Impuestos Internos
				'99	Otro	'Impuestos Provinciasles es Tributo id 2 	'Ingresos Brutos 
				AlicuotaTributo = 0 : BaseImponibleTributo = 0 : DescripcionTributo= "" : ImporteTributo = 0
			
				If esdelgada Then
					For Each oImpuesto In pComprobante.Impuestos
						If (oImpuesto.Importe > 0) Then
							If oImpuesto.DefinicionImpuesto.Impuesto.Codigo <> "010" Then	
								If mid(oImpuesto.DefinicionImpuesto.Impuesto.Tipo,1,3) = "PER" Then   'Perceiciones!!!!!
									If oImpuesto.Importe > 0 Then
										AlicuotaTributo 		 	= Round((Round(oImpuesto.Importe,2) * 100) / Round(pComprobante.SubTotal_Importe,2),2)
										BaseImponibleTributo 		= Round(pComprobante.SubTotal_Importe,2)
										DescripcionTributo 			= oImpuesto.DefinicionImpuesto.Impuesto.Nombre
  										importeTributo				= Round(oImpuesto.Importe,2)

'*******************************************************************************
'*******************************************************************************
                                                             '------ Agregado 08/03/2021 para Forzar Dolarización de Comprobantes
				                        If FuerzaDolarizacionComprobante Then
										   BaseImponibleTributo = Round((BaseImponibleTributo/CotizacionDolar),2)
										   importeTributo	= Round((importeTributo/CotizacionDolar),2)
                                        End If				
                                                              '------ Fin Agregado 08/03/2021 para Forzar Dolarización de Comprobantes
'*******************************************************************************
'*******************************************************************************

										If (oImpuesto.DefinicionImpuesto.Impuesto.SubTipo = "IVA") OR (oImpuesto.DefinicionImpuesto.Impuesto.SubTipo = "GAN") Then id_Tributo = 1 '1	Impuestos nacionales
										If oImpuesto.DefinicionImpuesto.Impuesto.SubTipo = "IIBB" Then id_Tributo = 2 '2	Impuestos provinciales

										ObtenerDatosComprobante = WSFACT.AgregarTributo(id_Tributo, DescripcionTributo, BaseImponibleTributo, AlicuotaTributo, importeTributo)
									End if
								End If
							End If
						End If
					Next
				else
					For Each oImpuesto In pComprobante.ImpuestosTransaccion
					   sCodigoImpuesto = oImpuesto.DefinicionImpuesto.Impuesto.Codigo												
					   If (oImpuesto.valor.Importe > 0) Then
'*******************************************************************************
'*******************************************************************************
'Esta linea es para Stenfar, y no se como Generalizarla sin hacer mucho codigo
'						  If (sCodigoImpuesto = "030") or (sCodigoImpuesto = "020") or (sCodigoImpuesto = "410") or (sCodigoImpuesto = "440") or (sCodigoImpuesto = "460")or (sCodigoImpuesto = "470") or (sCodigoImpuesto = "472") or (sCodigoImpuesto = "475") or (sCodigoImpuesto = "480") or (sCodigoImpuesto = "490") or (sCodigoImpuesto = "400") Then	
'Esta linea es para Stenfar, y no se como Generalizarla sin hacer mucho codigo

						  If oImpuesto.DefinicionImpuesto.Impuesto.Codigo <> "010" Then	
								If mid(oImpuesto.DefinicionImpuesto.Impuesto.Tipo,1,3) = "PER" Then   'Perceiciones!!!!!

									If oImpuesto.valor.Importe > 0 Then
										AlicuotaTributo 		= Round((Round(oImpuesto.valor.Importe,2) * 100) / Round(pComprobante.SubTotal.Importe,2),2)
										BaseImponibleTributo 	= Round(pComprobante.SubTotal.Importe,2)
										DescripcionTributo 	= oImpuesto.DefinicionImpuesto.Impuesto.Nombre  
										importeTributo 		= Round(oImpuesto.valor.Importe,2)
'*******************************************************************************
'*******************************************************************************
                                                             '------ Agregado 08/03/2021 para Forzar Dolarización de Comprobantes
				                        If FuerzaDolarizacionComprobante Then
										   BaseImponibleTributo = Round((BaseImponibleTributo/CotizacionDolar),2)
										   importeTributo	= Round((importeTributo/CotizacionDolar),2)
                                        End If				
                                                              '------ Fin Agregado 08/03/2021 para Forzar Dolarización de Comprobantes
'*******************************************************************************
'*******************************************************************************

										If (oImpuesto.DefinicionImpuesto.Impuesto.SubTipo = "IVA") OR (oImpuesto.DefinicionImpuesto.Impuesto.SubTipo = "GAN") Then id_Tributo = 1 '1	Impuestos nacionales
										If oImpuesto.DefinicionImpuesto.Impuesto.SubTipo = "IIBB" Then id_Tributo = 2 '2	Impuestos provinciales
										ObtenerDatosComprobante = WSFACT.AgregarTributo(id_Tributo, DescripcionTributo, BaseImponibleTributo, AlicuotaTributo, importeTributo)
									End if
								End If
							End If
						End If
					Next
				End if

			'------Activado el 1/7/2019 para MiPyMEs o Factura de Crédito
				If pComprobante.BOExtension.MiPyMEs Then
				    If tipo_cbte = 201 Then ' Solo FCE
                       If xWSMTXCA Then
                          ObtenerDatosComprobante = WSFACT.AgregarOpcional(21, sCbu, sAlias)  ' CBU (alias opcional) Solo Factura
 '------ Agregado 05/05/2021 por nueva validación en Comprobantes MiPyMEs
                          ObtenerDatosComprobante = WSFACT.AgregarOpcional(27, "SCA") ' ALIAS
 '------ Fin Agregado 05/05/2021 por nueva validación en Comprobantes MiPyMEs

                       Else
				   	      ObtenerDatosComprobante = WSFACT.AgregarOpcional(2101, sCbu)   ' CBU
				   	      ObtenerDatosComprobante = WSFACT.AgregarOpcional(2102, sAlias) ' ALIAS
'*******************************************************************************
'*******************************************************************************
                                                             '------ Agregado 01/04/2021 por nueva validación en Comprobantes MiPyMEs
				   	      ObtenerDatosComprobante = WSFACT.AgregarOpcional(27, "SCA") ' ALIAS
                                                             '------ Fin Agregado 01/04/2021 por nueva validación en Comprobantes MiPyMEs
'*******************************************************************************
'*******************************************************************************

					   End If
				    End if
				    If tipo_cbte = 203 Then 'Solo NC
						If pcomprobante.BOExtension.ANULAFACTURACOMPLETA Then
							ObtenerDatosComprobante = WSFACT.AgregarOpcional(22, "S")
						else
							ObtenerDatosComprobante = WSFACT.AgregarOpcional(22, "N")
						End if
				    End if
					If tipo_cbte = 202 Then 'Solo ND
						'Por ahora no lo usamos porque os debitos en afip no se estan haciendo, van todos comunes
					'	If pcomprobante.BOExtension.ANULAFACTURACOMPLETA Then
					'		ObtenerDatosComprobante = WSFACT.AgregarOpcional(22, "S")
					'	else
							ObtenerDatosComprobante = WSFACT.AgregarOpcional(22, "N")
					'	End if
				   End if
				   
				End If
			'------Fin Activado el 1/7/2019 para MiPyMEs o Factura de Crédito


'------ Agregado para WSMTXCA 10/10/2019
                If xWSMTXCA Then
					stop:stop
					For Each xItem In pComprobante.ItemsTransaccion
		               Set oPosicion = GetPosicionImpuestoPorCodigo(xItem.Referencia, "010")

		               If xItem.Referencia.BOExtension.UNIDADREFERENCIAAFIP = 0 Then
		                  u_mtx = 1
		               Else
		                  u_mtx = xItem.Referencia.BOExtension.UNIDADREFERENCIAAFIP
		               End If
                       cod_mtx = xItem.Referencia.BOExtension.CODIGOGTIN13
                       codigo = xItem.Referencia.Codigo
                       ds = xItem.Referencia.Descripcion
                       qty = xItem.Cantidad_Cantidad
                       umed = 7
		               If pComprobante.Nota = "A" Then
		                  precio = Round(xItem.Valor_Importe, 6)
		               Else
		                  precio = Round(xItem.Valor_Importe * (1 + (oPosicion.Coeficiente/100)), 6)
		               End If
                       bonif = "0.00"
                       cod_iva = 5
                       imp_iva = "21.00"
                       imp_subtotal = "121.00"

		               'Agrega el IVA:
						If (oPosicion.PosicionImpuesto.Codigo = "EXE") Then
							  'cod_iva = 3
							  cod_iva = 2
							  imp_iva = "0.00"		
							  imp_subtotal = xItem.Total_Importe
						Else
							  Set oIva = GetImpuestoTdPorCodigo(xItem, "010")
							  xTotalItem = xItem.Total_Importe
							  xIVAItem = oIva.Importe
							  If (oPosicion.Coeficiente = 0) Then
								 cod_iva = 3
							  ElseIf (oPosicion.Coeficiente = 10.5) Then
								 cod_iva = 4
							  ElseIf (oPosicion.Coeficiente = 21) Then
								 cod_iva = 5
							  ElseIf (oPosicion.Coeficiente = 27) Then
								 cod_iva = 6
							  End If
							  If pComprobante.Nota = "A" Then
								 imp_iva = xIVAItem		
								 imp_subtotal = xTotalItem + xIVAItem
							  Else
								 imp_iva = "0.00"		
								 imp_subtotal = xTotalItem + xIVAItem
							  End If
						End If
stop					   

'*******************************************************************************
'*******************************************************************************
                       '------ Agregado 08/03/2021 para Forzar Dolarización de Comprobantes
						If FuerzaDolarizacionComprobante Then
							If precio > 0 Then
								precio = Round((precio/CotizacionDolar),2)
							End If
							If bonif > 0 Then
								 bonif = Round((bonif/CotizacionDolar),2)
							End If
							If imp_iva > 0 Then
								imp_iva = Round((imp_iva/CotizacionDolar),2)
							End If
							If imp_subtotal > 0 Then
								imp_subtotal = Round((imp_subtotal/CotizacionDolar),2)
							End If
						End If
                       '------ Fin Agregado 08/03/2021 para Forzar Dolarización de Comprobantes
'*******************************************************************************
'*******************************************************************************

                       ok = WSFACT.AgregarItem(u_mtx, cod_mtx, codigo, ds, qty, umed, precio, bonif, cod_iva, imp_iva, imp_subtotal)

					Next
                End If
				
'------ FIN Agregado para WSMTXCA 10/10/2019


End Function
			
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' @funcion ObtenerMoneda
' @descripcion Obtengo el codigo de Moneda en funcion a la tabla moneda de calipso
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Private Function ObtenerMoneda(ByRef pComprobante)
	
   If pComprobante.TipoTransaccion.TDDEF Is Nothing Then
	  If (Not pComprobante.Total.UnidadValorizacion Is Nothing) Then
	     If (pComprobante.Total.UnidadValorizacion.Id = cPesos) Then
		    ObtenerMoneda = "PES"
		 ElseIf (pComprobante.Total.UnidadValorizacion.Id = cDolares) Then
			ObtenerMoneda = "DOL"
		 ElseIf (pComprobante.Total.UnidadValorizacion.Id = cReales) Then
			ObtenerMoneda = "012"
		 ElseIf (pComprobante.Total.UnidadValorizacion.Id = cEuros) Then
			ObtenerMoneda = "060"
		End If
	  End If
   Else
	  If Not pComprobante.Total_UnidadValorizacion Is Nothing Then
	   	 If (pComprobante.Total_UnidadValorizacion.Id = cPesos) Then
		    ObtenerMoneda = "PES"
		 ElseIf (pComprobante.Total_UnidadValorizacion.Id = cDolares) Then
			ObtenerMoneda = "DOL"
		 ElseIf (pComprobante.Total_UnidadValorizacion.Id = cReales) Then
			ObtenerMoneda = "012"
		 ElseIf (pComprobante.Total_UnidadValorizacion.Id = cEuros) Then
			ObtenerMoneda = "060"	
		 End If
	  End If
   End If
End Function


''Solo Stenfar

Function ObtieneTransicion ( Transaccion)    
		ObtieneTransicion = ""
	    xOtp = Transaccion.Tipotransaccion.OTP
		If xOtp = "TRFACTURAVENTA" then
		 	set o_Transicion  = instanciarbo("23B6CDFB-B6A0-4FE0-803A-73D42D403E2D", "TRANSICION", Transaccion.workspace) 'Transicion para Cerrar
			registraractividadplaunch("Ejecutando Transicion")
		 End if
		 If xOtp = "TRCREDITOVENTA" then
		 	If Transaccion.Tipotransaccion.id = "814E5B92-A616-11D5-B2B1-00D0B7BFE069" then
			   set o_Transicion  = instanciarbo("1A1B60B2-A92F-4FF9-A79A-0B390AEDE69D", "TRANSICION", Transaccion.workspace) 'Transicion para Cerrar
    			registraractividadplaunch("Ejecutando Transicion")
			else
		 		set o_Transicion  = instanciarbo("CC0F7113-2EB8-4951-BB52-186E2DBE8B7B", "TRANSICION", Transaccion.workspace) 'Transicion para Cerrar
		    End if

		 End if
		 If xOtp = "TRDEBITOVENTA" then
		 	set o_Transicion  = instanciarbo("B9E10404-A9F4-4053-AD45-BA01393628C1", "TRANSICION", Transaccion.workspace) 'Transicion para Cerrar
			registraractividadplaunch("Ejecutando Transicion")
		 End if

		 If not o_Transicion is nothing then
		 	ObtieneTransicion = o_Transicion.Descripcion
		 End if
End Function
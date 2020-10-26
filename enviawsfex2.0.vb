'' @script EnviaWsfex 2.0   07-10-2020
'EnviaWsfe  Con WSFE1 y WSMTXCA
'descripcion Autoriza el comprobante conectandose al servicio web de la AFIP con detalle de Productos.
'-------------------------------------------------------------------------------
Public AlicuotaTributo,	BaseImponibleTributo,DescripcionTributo,ImporteTributo,aToken,aSign,nro_calipso_Actual,id_Solicitud,ExisteSolicitud,sCbu,sAlias,_
	   nContribuyente,imp_Tolerancia,cPesos,cDolares,cEuros,cReales,xValid,xSintecrom,xZSA,xCabal,xEgramar,xProdalsa,xBinova,xWSMTXCA,xUtilizaPathInstalacionSA,CAE,xStenfar



Private gInteractivo


Sub Main()
Stop
    
'*********************************************************************************************************************************************
	'Variables Modificables segun el Cliente
'*********************************************************************************************************************************************
	'Tolerancia de Centavos en Total del Comproabnte
	imp_Tolerancia 	= 0.02  
	

'Path del Certificado
'Si le asignamos True, usa el Path de instalación de Sistemas Agiles. Ojo!!! El problema de esto es el mantenimiento, porque hay que modificarlo en todas las maquinas
'Si le asignamos False, utiliza directamente los que están en la red y que tenemos que asignar en los parametros "FACTURAELECTRONICA_CERTIFICADO" y "FACTURAELECTRONICA_CLAVE" con el path completo. Ejemplo: Z:\FE\calipso_24e52a790ad63ee0.crt o \\SQLSERVER\Util\fe\calipso---.crt 
xUtilizaPathInstalacionSA = False
	
'Empresa
xValid     = False
xSintecrom = False
xZSA       = False
xCabal     = False
xEgramar   = False
xProdalsa  = False
xBinova    = False
xStenfar   = True

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
	cPesos 			= "{13207D68-8854-11D5-B08A-004854841C8A}"
	cDolares 		= "{13207D6B-8854-11D5-B08A-004854841C8A}"
	cReales			= "{D8C0F91C-49C0-463C-A68D-26CD5FB7DC05}"
	CEuros			= "{D783CBDE-FA11-4B0A-8590-0018AF050E3C}"
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


	
'*********************************************************************************************************************************************
'*********************************************************************************************************************************************	
	aToken			= ""
	aSign			= ""
	
	ausuario		= nombreusuario()
	aComputername	= GetComputerName(  )
	aNuevoTicketAccesso = True
	
	set pComprobante = aComprobante.Value
	gInteractivo	 = aInteractivo.Value
	
'Obtengo Parametros 	
	Certificado  		= ExisteBO( pComprobante, "ASOCIACION", "Clave", "FACTURAELECTRONICA_CERTIFICADO", nil, FALSE, FALSE, "=" ).valor
	ClavePrivada  	= ExisteBO( pComprobante, "ASOCIACION", "Clave", "FACTURAELECTRONICA_CLAVE", nil, FALSE, FALSE, "=" ).valor
	xPathXml			= ExisteBO( pComprobante, "ASOCIACION", "Clave", "FACTURAELECTRONICA_PATHLOG", nil, FALSE, FALSE, "=" ).valor
	nContribuyente	= ExisteBO( pComprobante, "ASOCIACION", "Clave", "FACTURAELECTRONICA_CUIT", nil, FALSE, FALSE, "=" ).valor
	sCbu				= ExisteBO( pComprobante, "ASOCIACION", "Clave", "MIPYMES_CBU", nil, FALSE, FALSE, "=" ).valor
	sAlias 			= ExisteBO( pComprobante, "ASOCIACION", "Clave", "MIPYMES_ALIAS", nil, FALSE, FALSE, "=" ).valor
	
	If xWSMTXCA Then
       sServicio = "wsmtxca"
    Else
       sServicio = "wsfex"
    End If
	
	xTicketAcceso	= False
	
	xTicketAcceso = ObtenerTicketAccesso (pComprobante,nContribuyente,sServicio,aComputername,ausuario,Certificado,ClavePrivada)
	
	
	
'30  Crear objeto interface Web Service de Factura Electrónica de Mercado Interno

	If xTicketAcceso Then 
		
		set WSFACT		= CreaWSFACT(aToken,aSign,nContribuyente)
   
'40 Establezco los valores de la factura a autorizar:

        If xEgramar Then
		   punto_vta = CInt(pComprobante.puntoventa.codigo)
		   'punto_vta = string((4-len(punto_vta)),"0")+punto_vta
        Else
			If xCabal Then
				punto_vta 		=  CInt(pComprobante.NUMERADOR.NUMERADOR.CARACTERESPREFIJO)
			Else
			   punto_vta = CInt(pComprobante.BOExtension.PuntoVenta.Codigo)
			End if
        End If
				
		esdelgada		= ObtenerTrDelada(pComprobante)
		
		tipo_cbte 		= ObtenerTipoComprobante(pComprobante)
		
		cbte_nro_calipso = "" : cbte_nro_afip = ""
		
		cbte_nro_calipso = CLng(ObtenerUltimoNumeroCalipso (pComprobante,tipo_cbte,punto_vta))
	
		cbte_nro_afip 	= WSFACT.GetLastCMP(tipo_cbte, punto_vta)
		
		ControlarExcepcion WSFACT
		

		senddebug WSFACT.errmsg : senddebug WSFACT.errcode
		
		If cbte_nro_afip = "" Then
			cbte_nro_afip = 0                ' no hay comprobantes emitidos
		Else
			cbte_nro_afip = CLng(cbte_nro_afip)   ' convertir a entero largo
		End If
		
	
		ExisteSolicitud = False
		nro_calipso_Actual = 0
		
		call ObtenerUltimaSolicitud(pComprobante)  'Obtiene id_Solicitud,nro_calipso_Actual,ExisteSolicitud
		

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
				
				id = CStr(CCur(WSFACT.GetLastID()) + 1)
				
				CAE = WSFACT.Authorize(id)
				ControlarExcepcion WSFACT
				' Imprimo pedido y respuesta XML para depuración (errores de formato)		
				senddebug "Resultado"& WSFACT.Resultado : senddebug "CAE"& WSFACT.CAE :	senddebug "Numero de comprobante:"& WSFACT.CbteNro
				senddebug WSFACT.XmlRequest : senddebug WSFACT.XmlResponse :	senddebug "Reproceso:"& WSFACT.Reproceso 
			' Muestro los errores
				amensaje = ""
				If WSFACT.errmsg <> "" Then amensaje = replace(WSFACT.errmsg,chr(10),"")

				' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)

				For Each evento In WSFACT.eventos:
					Call CustomMsgBox(evento, vbInformation, "Factura Electrónica")
				Next
	
				If WSFACT.Resultado = "A" and CAE <> "" Then 
					
					x = ActualizaSolicitud (pComprobante,id_Solicitud,amensaje,WSFACT.Resultado)
					
					x = ActualizaCamposTransaccion(pcomprobante,WSFACT,punto_vta,tipo_cbte)
	
					x = ActualizaComprobante (pComprobante,WSFACT.Vencimiento,tipo_cbte,punto_vta,WSFACT.CbteNro,WSFACT.FechaCbte)
					
					senddebug ("Numero: " & CStr(punto_vta ) & Right("00000000" & CStr(cbte_nro), 8)):senddebug ("Cae   : "& WSFACT.CAE)
					EVento2 = "Resultado:" & WSFACT.Resultado & " CAE: " & CAE & " Venc: " & WSFACT.Vencimiento & " Obs: " & WSFACT.obs & " Reproceso: " & WSFACT.Reproceso
					registraractividadplaunch( "Procesando en Afip" )
					
					ReturnValue = True

				Else
				   If amensaje = "" Then amensaje = replace(WSFACT.obs,"'","")
				   
				   x = ActualizaSolicitud (pComprobante,id_Solicitud,amensaje,WSFACT.Resultado)
				   
					xtext = WSFACT.XmlResponse
					xNombreArchivoXML = nContribuyente & CStr(punto_vta ) & Right("00000000" & CStr(cbte_nro), 8) &  strFechaCae & id_Solicitud&".xml"
					xArchivoXML = xPathXml & xNombreArchivoXML
					call EscribirTXT( xArchivoXML, xtext, 0 ) 
					EVento2 = "Resultado:" & WSFACT.Resultado & " CAE: " & CAE & " Venc: " & WSFACT.Vencimiento & " Obs: " & WSFACT.obs & " Reproceso: " & WSFACT.Reproceso
					ReturnValue = False
					Call CustomMsgBox(EVento2, vbInformation + vbOKOnly, "Factura Electrónica")
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
				cae2 = WSFACT.GetCMP(tipo_cbte, punto_vta, cbte_nro)
				ControlarExcepcion WSFACT
				
				ok = WSFACT.AnalizarXml("XmlResponse")
				
				x = ActualizaCamposTransaccion(pcomprobante,WSFACT,punto_vta,tipo_cbte)
							
				x = ActualizaComprobante (pComprobante,WSFACT.Vencimiento,tipo_cbte,punto_vta,cbte_nro,WSFACT.FechaCbte)
				ReturnValue = True				   
				senddebug "Fecha Comprobante:"& WSFACT.FechaCbte : senddebug "Fecha Vencimiento CAE"& WSFACT.Vencimiento :senddebug "Importe Total:"& WSFACT.ImpTotal :	senddebug "Resultado:"& WSFACT.Resultado			
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
			tra = WSAA.CreateTRA(sServicio, ttl)
			ControlarExcepcion WSAA
			senddebug tra
			
			' Generar el mensaje firmado (CMS)
		    If xUtilizaPathInstalacionSA Then
			   Path = WSAA.InstallDir + "\conf\"    ' Especificar la ubicacion de los archivos certificado y clave privada
			   cms  = WSAA.SignTRA(tra, Path + Certificado, Path + ClavePrivada)
            Else		
			   cms = WSAA.SignTRA(tra, Certificado, ClavePrivada)
			End If

			ControlarExcepcion WSAA
			senddebug cms
			' Conectarse con el webservice de autenticación:
			cache 	= ""
			proxy   = ""  '172.16.1.26:8080
			wrapper = "" ' libreria http (httplib2, urllib2, pycurl)

			cacert 	= WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante
			cacert 	= ""
			wsdl  	= "https://wsaa.afip.gov.ar/ws/services/LoginCms?wsdl"

			ObtenerTicketAccesso = WSAA.Conectar(cache, wsdl, proxy, wrapper, cacert) ' Homologación
			ControlarExcepcion WSAA
			' Llamar al web service para autenticar:
			ta = WSAA.LoginCMS(cms)
			ControlarExcepcion WSAA
			
			'x = InsertSQL( "WSAA_TICKETACCESO", "('"&nContribuyente&"','"&sServicio&"',GETDATE(),'"&WSAA.Sign&"','"&WSAA.Token&"',0000000000,DATEADD(N, 600, GETDATE()),'"&aComputername&"','"&ausuario&"','4.0.0.0',9999)", self.WorkSpace )
			aQueryInsert = "INSERT INTO [WSAA_TICKETACCESO] ([CONTRIBUYENTE], [SERVICIO], [GENERACION], [SIGN], [TOKEN], [UNIQUEID], [VENCIMIENTO], [EQUIPO], [USUARIO], [VERSION], [PROCESO]) VALUES ('"&nContribuyente&"','"&sServicio&"',GETDATE(),'"&WSAA.Sign&"','"&WSAA.Token&"',0000000000,DATEADD(N, 600, GETDATE()),'"&aComputername&"','"&ausuario&"','4.0.0.0',9999)"
			 Set xRst = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aQueryInsert )
			 aToken = WSAA.Token
			 aSign	= WSAA.Sign

			 cacert = WSAA.InstallDir & "\conf\afip_ca_info.crt" ' certificado de la autoridad de certificante (solo pycurl)
			 cacert 	= ""

		End if    
	
	End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' @funcion CreaWSFACT - reemplaza a CreaWSFEv1 o CreaWSMTXCA y ahora depende del parametro de la empresa xWSMTXCA 
' @descripcion Conecta con WSFEv1 o WSMTXCA de sistemas Agiles
'-------------------------------------------------------------------------------	
Public Function CreaWSFACT(aToken,aSign,nContribuyente)
	
		proxy 			= "" ' "usuario:clave@localhost:8000"
		wsdl 			= "https://servicios1.afip.gov.ar/wsfexv1/service.asmx?WSDL"
		cache 			= "" 'Path
		wrapper 		= "" ' libreria http (httplib2, urllib2, pycurl)
	
		Set CreaWSFACT = CreateObject("WSFEXv1")
		senddebug CreaWSFACT.Version
		CreaWSFACT.Token 	= aToken
		CreaWSFACT.Sign  	= aSign

		CreaWSFACT.Cuit = nContribuyente     'Cuit de Cabal
		' deshabilito errores no manejados
		CreaWSFACT.LanzarExcepciones = False
		' Conectar al Servicio Web de Facturación
   
		ok = CreaWSFACT.Conectar(cache, wsdl, proxy, wrapper, cacert) ' homologación
		ControlarExcepcion CreaWSFACT
		' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
		CreaWSFACT.Dummy
		ControlarExcepcion CreaWSFACT
	End Function
		
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	
Sub ControlarExcepcion(Obj)
    ' Nueva funcion para verificar que no haya habido errores:
    On Error GoTo 0

    If obj.Excepcion <> "" Then
        ' Depuración (grabar a un archivo los detalles del error)
  	   Call CustomMsgBox(obj.Excepcion, vbExclamation, "Factura Electrónica - Excepción")
    End If

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function	Inserta_Wsfe_Comprobante (pComprobante,tipo_cbte,punto_vta)
				

	
    aqueryInsertComprobante = "INSERT INTO [WSFEX_COMPROBANTE] ([ID], [TIPO], [PUNTOVENTA], [NUMERO], [CAE], [VENCIMIENTOCAE]) VALUES ('"&pcomprobante.id&"',"& tipo_cbte&","&punto_vta&",0,'','')"
	Set xRst = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aqueryInsertComprobante )
	Inserta_Wsfe_Comprobante = True
End Function				

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function	Inserta_wsfe_Solicitud (pComprobante,nContribuyente,aComputername,ausuario,Pendiente,aMensaje)
	aPendiente = 0
	aError	   = 0
	If Pendiente Then aPendiente = 1
	If amensaje <> "" Then aError = 1
	aqueryInsertSolicitud = "INSERT INTO [WSFEX_SOLICITUD] ([CONTRIBUYENTE], [COMPROBANTE_ID], [PENDIENTE],[ERROR],[MENSAJE],[MOMENTO], [EQUIPO], [USUARIO], [VERSION]) VALUES ('"&nContribuyente&"','"&pcomprobante.id&"',"&aPendiente&","&aError&",'"&amensaje&"',GETDATE(),'"&aComputername&"','"&ausuario&"','4.0.0.0')"
	
	Set xRst = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aqueryInsertSolicitud )
	Inserta_wsfe_Solicitud = True
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function	ObtenerUltimaSolicitud(pComprobante)
		
		aQueryBuscarUltimaSolicitud = "SELECT TOP 1 [A].ID,ISNULL([B].NUMERO,0) AS NUMERO FROM [WSFEX_SOLICITUD] [A] JOIN [WSFEX_COMPROBANTE] [B] ON ([B].[ID] = [A].[COMPROBANTE_ID]) WHERE ([A].[COMPROBANTE_ID] = '"&pcomprobante.id&"') ORDER BY [A].[ID] DESC"
		Set xRstSol = RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aQueryBuscarUltimaSolicitud )
		Do While Not xRstSol.EOF
			ExisteSolicitud 	= True
			id_Solicitud     	= xRstSol("ID").Value
			nro_calipso_Actual 	= xRstSol("NUMERO").Value
			xRstSol.MoveNext
		Loop	
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ObtenerUltimoNumeroCalipso (pComprobante,tipo_cbte,punto_vta)
		ObtenerUltimoNumeroCalipso = 0
		aQueryUltimoNro = "select top 1 NUMERO from WSFEX_COMPROBANTE where TIPO = "&tipo_cbte&" and PUNTOVENTA = "&punto_vta &" order by NUMERO desc"
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
		aqueryUpdateSolicitud	="UPDATE [WSFEX_SOLICITUD] SET [PENDIENTE] = 0, [ERROR] = 0, [MENSAJE] = '' WHERE ([ID] = "&id_Solicitud&")"
	else
		aqueryUpdateSolicitud	="UPDATE [WSFEX_SOLICITUD] SET [PENDIENTE] = 0, [ERROR] = 1, [MENSAJE] = '"&amensaje&"' WHERE ([ID] = "&id_Solicitud&")"
	End if
	
	Set xRst 				= RecordSet( StringConexion( "Conexion_ADO", pComprobante.WorkSpace ), aqueryUpdateSolicitud )
	ActualizaSolicitud = True

End Function	
	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function	ActualizaComprobante (pComprobante,Vencimiento,tipo_cbte,punto_vta,cbte_nro,fecha_cbte)
		ActualizaComprobante = False
		Vencimiento			 = mid(Vencimiento,7,4)&mid(Vencimiento,4,2)&mid(Vencimiento,1,2)
		
		aqueryInsertComprobante = " IF EXISTS(SELECT [ID] FROM [WSFEX_COMPROBANTE] WHERE ([ID] = '"&pcomprobante.id&"'))"&_
							   " BEGIN "&_
							   " UPDATE [WSFEX_COMPROBANTE] SET [NUMERO] = "&cbte_nro&", [TIPO] = "&tipo_cbte&", [CAE] = '"&pComprobante.BOExtension.CAE&"', [VENCIMIENTOCAE] = '"&Vencimiento&"' WHERE ([ID] = '"&pcomprobante.id&"');"&_
							   " END; "&_
							   " ELSE "&_
							   " BEGIN "&_
							   " INSERT INTO [WSFEX_COMPROBANTE] ([ID], [TIPO], [PUNTOVENTA], [NUMERO], [CAE], [VENCIMIENTOCAE] ) VALUES ('"&pcomprobante.id&"',"& tipo_cbte&","&punto_vta&","&cbte_nro&",'"&pComprobante.BOExtension.CAE&"','"&Vencimiento&"');"&_
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
		   pComprobante.BOExtension.FECHAVTOCAE = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.BOExtension.FECHAVTOCAE) & Right("0" & Month(pComprobante.BOExtension.FECHAVTOCAE), 2) & Right("0" & Day(pComprobante.BOExtension.FECHAVTOCAE), 2)
		   pComprobante.BOExtension.CODIGODEBARRAS = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
        End If
					 
'***************************************************************************************************************************************	
'Sintecrom
        If xSintecrom Then
		   pComprobante.BOExtension.CAE	= WSFACT.CAE
		   pComprobante.BOExtension.FECHAVTOCAE = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.BOExtension.FECHAVTOCAE) & Right("0" & Month(pComprobante.BOExtension.FECHAVTOCAE), 2) & Right("0" & Day(pComprobante.BOExtension.FECHAVTOCAE), 2)
		   pComprobante.BOExtension.CODIGODEBARRAS = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
        End If

'***************************************************************************************************************************************	
'Cabal
        If xCabal Then
		   pComprobante.BOExtension.CAE	= WSFACT.CAE
		   pComprobante.BOExtension.FECHAVTOCAE = cdate(mid(WSFACT.Vencimiento,7,2)&"/"&mid(WSFACT.Vencimiento,5,2)&"/"&mid(WSFACT.Vencimiento,1,4))
		   strFechaCae = Year(pComprobante.BOExtension.FECHAVTOCAE) & Right("0" & Month(pComprobante.BOExtension.FECHAVTOCAE), 2) & Right("0" & Day(pComprobante.BOExtension.FECHAVTOCAE), 2)
		   pComprobante.BOExtension.CODIGODEBARRAS = ArmoCodigoBarra(pComprobante,tipo_cbte,WSFACT.Cuit,Right("00000" & CStr(punto_vta ),5),WSFACT.CAE,strFechaCae)
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


		ActualizaCamposTransaccion = True

				
End Function


					 
' @funcion ObtenerTipoComprobante
' @descripcion Obtiene el código de tipo de comprobante.
'-------------------------------------------------------------------------------
Private Function ObtenerTipoComprobante(ByRef pComprobante)

	Dim iTipoComprobante
	
	' Establece el tipo de comprobante en base a la letra de la factura.
	
		' Establece el tipo de comprobante de acuerdo al tipo de factura.
	sClassName = ClassName(pComprobante)
	If (sClassName = "TRFACTURAVENTA") Or (sClassName = "TDFACTURAVENTA") Then
	
	    iTipoComprobante = 19

	ElseIf (sClassName = "TRDEBITOVENTA") Or (sClassName = "TDDEBITOVENTA") Then 
	
	    iTipoComprobante = 20
	
	ElseIf (sClassName = "TRCREDITOVENTA") Or (sClassName = "TDCREDITOVENTA") Then
	
	    iTipoComprobante = 21
	
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

           If xZSA Then
			  dVencimiento = dVencimiento + oItemTipoPago.DiasVencimiento 
           Else
			  dVencimiento = dVencimiento + oItemTipoPago.Periodo 
           End If
		    
			Exit For
		Next
    Else
	        dVencimiento = dVencimiento + 2
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
   
      xSQL = "SELECT calipso.FN_CODIGOBARRA_CAE('" & strcuit & strcodigocomprobante & strPuntoVenta & strCAE & strFecha & "') AS CODIGO "
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
	If (sClassName = "TDCREDITOVENTA") Or (sClassName = "TDFACTURAVENTA") Or (sClassName = "TDDEBITOVENTA") Then
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
Private Function ObtenerDatosComprobante ( WSFEX, pComprobante,cbte_nro_calipso,esdelgada,tipo_cbte,punto_vta)
				ObtenerDatosComprobante = False
				
				fecha = year(pComprobante.FechaActual) & right("00" & month(pComprobante.FechaActual),2) & right("00" & day(pComprobante.FechaActual),2)

				concepto = 1  'Bienes

				tipo_doc = 80
				nro_doc = pComprobante.Destinatario.EnteAsociado.Cuit
				nro_doc = replace(nro_doc,"-","")
				
				For Each oItem In pComprobante.ItemsTransaccion
    				tipo_expo	= 1	
					If classname(oItem.Referencia) = "PRODUCTO" Then
		   			   tipo_expo = 1 'Exportación definitiva de bienes
					Else
						' Si es una Factura y de serivicios
		   				If tipo_cbte = 19 then tipo_expo	= 2	'Servicios
				    end if 
					Exit For
				Next
				If tipo_cbte = 19 Then
				   If tipo_expo = 2 then
				   	  permiso_existente = ""
				   Else 
					  permiso_existente = "N" 'Indica si se posee documento aduanero de exportación (permiso de embarque)
				   End if
				Else
					permiso_existente = ""
				End If
				If Not pComprobante.DomicilioEntrega.Pais Is Nothing Then
					dst_cmp = pComprobante.DomicilioEntrega.Pais.codigo
				End If
				cliente				= pComprobante.Destinatario.Denominacion
				cuit_pais_cliente 	= Replace(pComprobante.Destinatario.EnteAsociado.CUIT, "-", "")
				domicilio_cliente 	= pComprobante.DomicilioEntrega.Calle
				id_impositivo  	  	= Replace(pComprobante.Destinatario.EnteAsociado.CUIT, "-", "")
				obs_comerciales	  	= ""
				obs 			  	= ""
				forma_pago		  	= pComprobante.TipoPago.Observacion
				incoterms		  	= "EXW"
				cbte_nro 		  	= cbte_nro_calipso + 1
				fecha_cbte 			= fecha
				idioma_cbte 		= 1 
				moneda_id 			= ObtenerMoneda(pComprobante) : moneda_ctz = pComprobante.Cotizacion
				imp_total 			=  Round(pComprobante.ValorTotal,2)
				incoterms_ds		= ""
				fecha_pago 			= ObtenerVencimientoPago(pComprobante)
				If tipo_expo = 1 Then
				ok = WSFEX.CrearFactura(tipo_cbte, punto_vta, cbte_nro, fecha_cbte, _
					imp_total, tipo_expo, permiso_existente, dst_cmp, _
					cliente, cuit_pais_cliente, domicilio_cliente, _
					id_impositivo, moneda_id, moneda_ctz, _
					obs_comerciales, obs, forma_pago, incoterms, _
					idioma_cbte,incoterms_ds)
				else
				ok = WSFEX.CrearFactura(tipo_cbte, punto_vta, cbte_nro, fecha_cbte, _
					imp_total, tipo_expo, permiso_existente, dst_cmp, _
					cliente, cuit_pais_cliente, domicilio_cliente, _
					id_impositivo, moneda_id, moneda_ctz, _
					obs_comerciales, obs, forma_pago, incoterms, _
					idioma_cbte,incoterms_ds,fecha_pago)
				End if
				For Each oItemComprobante In pComprobante.ItemsTransaccion

					acodigo  		= oItemComprobante.referencia.codigo
					adescripcion 	= oItemComprobante.descripcion
					apreciounitario = Round((oItemComprobante.valor.importe) ,3)
					acantidad 		= oItemComprobante.cantidad.cantidad
					abonificacion   = Round((oItemComprobante.importebonificado),3)
					atotal 		    = Round((oItemComprobante.total.importe),3)
					aunidadmedida 	= 7 'Unidades	
					ok = WSFEX.AgregarItem(acodigo, adescripcion, acantidad, aunidadmedida, apreciounitario, atotal)
				
				Next

			
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

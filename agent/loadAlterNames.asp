<%
set ra = Server.CreateObject("ADODB.RecordSet")
If Request("LoadLanID") <> "" Then Session("LanID") = Request("LoadLanID")
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetAlterNames" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LanID") = Session("LanID")
set ra = cmd.execute()

txtAgent = ra("Singular")
txtAgents = ra("Plural")

ra.movenext
txtClient = ra("Singular")
txtClients = ra("Plural")

ra.movenext		'Cotizacion
txtQuote = ra("Singular")
txtQuotes = ra("Plural")

ra.movenext
txtPoll = ra("Singular")
txtPolls = ra("Plural")

ra.movenext
txtCXC = ra("Singular")

ra.movenext		'Facturas
txtInv = ra("Singular")
txtInvs = ra("Plural")

ra.movenext
txtNews = ra("Singular")
txtNewss = ra("Plural")

ra.movenext		'Ordenes de Venta
txtOrdr = ra("Singular")
txtOrdrs = ra("Plural")

ra.movenext
txtProm = ra("Singular")
txtProms = ra("Plural")

ra.movenext		'Recibos
txtRct = ra("Singular")
txtRcts = ra("Plural")

ra.movenext
txtOfert = ra("Singular")
txtOferts = ra("Plural")

ra.movenext		'Entregas
txtOdln = ra("Singular")
txtOdlns = ra("Plural")

ra.movenext		'Devoluciones en Vneta
txtOrdn = ra("Singular")
txtOrnds = ra("Plural")

ra.movenext		'Nota de credito deudores
txtOrin = ra("Singular")
txtOrins = ra("Plural")

ra.movenext		'Orden de compra
txtOpor = ra("Singular")
txtOpors = ra("Plural")

ra.movenext		'Entrada de mercancias OP (Consignacion)
txtOpdn = ra("Singular")
txtOpdns = ra("Plural")

ra.movenext		'Devolucion de Compra
txtOrpd = ra("Singular")
txtOrpds = ra("Plural")

ra.movenext		'Factura de compra
txtOpch = ra("Singular")
txtOpchs = ra("Plural")

ra.movenext		'Nota de credito acreedores
txtOrpc = ra("Singular")
txtOrpcs = ra("Plural")

ra.movenext		'Pagos a proveedor
txtOvpm = ra("Singular")
txtOvpms = ra("Plural")

ra.movenext
txtTax = ra("Singular")

ra.movenext
If userType = "C" Then
	txtBasketMinRep = ra("Singular")
Else
	txtBasketMinRep = ra("Plural")
End If

ra.movenext
If userType = "C" Then
	txtRef2 = ra("Singular")
Else
	txtRef2 = ra("Plural")
End If

ra.movenext
If userType = "C" Then
	txtAlterGrp = ra("Singular")
Else
	txtAlterGrp = ra("Plural")
End If

ra.movenext
If userType = "C" Then
	txtAlterFrm = ra("Singular")
Else
	txtAlterFrm = ra("Plural")
End If


ra.movenext		'Facturas Reservadas
txtInvRes = ra("Singular")
txtInvsRes = ra("Plural")

ra.movenext		'Solicitud anticipo
txtODPIReq = ra("Singular")
txtODPIReqs = ra("Plural")

ra.movenext		'Factura anticipo
txtODPIInv = ra("Singular")
txtODPIInvs = ra("Plural")






%>
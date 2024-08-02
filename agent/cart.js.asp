<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<!--#include file="authorizationClass.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Dim myAut
set myAut = New clsAuthorization

%>
<!--#include file="chkLogin.asp" -->
<!--#include file="lang/cart.js.asp" -->
<!--#include file="myHTMLEncode.asp"-->
//Variables generales
var SaveImgField;
var SaveImgImage;
var SaveImgMaxSize;
var vSetLineDisc = true;
var changePrice = true;
var searchCmd = 'd';

//Cantidad de decimales segun formato
var RateDec = <%=myApp.RateDec%>; 
var SumDec = <%=myApp.SumDec%>; 
var PriceDec = <%=myApp.PriceDec%>; 
var QtyDec = <%=myApp.QtyDec%>; 
var PercentDec = <%=myApp.PercentDec%>; 
var MeasureDec = <%=myApp.MeasureDec%>;
var Inc3dx = 0;
var IncBtch = 0;
var Verfy3dxOrder = '<%=myApp.Verfy3dxOrder%>';
var VerfyBtchOrder = '<%=myApp.VerfyBtchOrder%>';
var txtVolDiscount = '<%=getcartjsLngStr("LtxtVolDiscount")%>';
<% If userType = "C" Then %>
var lblItemDetailsQty = '<%=getcartjsLngStr("DtxtQty")%>';
var lblItemDetailsPrice = '<%=getcartjsLngStr("DtxtPrice")%>';
<% End If %>

<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKCartHasDynData" & Session("ID")
cmd.Parameters.Refresh
cmd("@UserType") = userType
set rs = Server.CreateObject("ADODB.RecordSet")
set rs = cmd.execute()
%>
var hasDynamicCols = <%=rs(0)%>;

function chkThis(Field, FType, EditType, FSize)
{
	switch(FType)
	{
		case "A":
			if (Field.value.length > FSize)
			{
				alert("<%=getcartjsLngStr("DtxtValFldMaxChar")%>".replace("{0}", FSize));
				Field.value = Left(Field.value,FSize);
			}
			break;
		case "N":
			switch (EditType)
			{
				case "":
				case " ":
					if (!MyIsNumeric(getNumericVB(Field.value)) && Field.value != "")
					{
						Field.value = "";
						alert("<%=getcartjsLngStr("DtxtValNumVal")%>");
					}
					else if (parseFloat(getNumericVB(Field.value))-parseInt(getNumericVB(Field.value)) != 0)
					{
						Field.value = "";
						alert("<%=getcartjsLngStr("DtxtValNumValWhole")%>");
					}
					break;
			}
			break;
		case "B":
			if (!MyIsNumeric(getNumericVB(Field.value)) && Field.value != "")
			{
				Field.value = "";
				alert("<%=getcartjsLngStr("DtxtValNumVal")%>");
			}
			else 
			{
				switch (EditType)
				{
					case "R":
						Field.value = OLKFormatNumber(getNumericVB(Field.value), RateDec);
						break;
					case "S":
						Field.value = OLKFormatNumber(getNumericVB(Field.value), SumDec);
						break;
					case "P":
						Field.value = OLKFormatNumber(getNumericVB(Field.value), PriceDec);
						break;
					case "Q":
						Field.value = OLKFormatNumber(getNumericVB(Field.value), myApp.QtyDec);
						break;
					case "%":
						Field.value = OLKFormatNumber(getNumericVB(Field.value), PercentDec);
						break;
					case "M":
						Field.value = OLKFormatNumber(getNumericVB(Field.value), MeasureDec);
						break;
				}
			}
			break;
	}
}


function doProcGroupNum(groupNum, loadTotals)
{
	$.post("cart/cartProcess.asp?d=" + (new Date()).toString(), { Field: 'GroupNum', FieldType: 'N', Value: groupNum, LineID: -1 },
	   function(data){
	   		var arrData = data.split('{S}');
	   		document.frmCart.DocDueDate.value = arrData[0];
	   		if (loadTotals)
	   		{
					document.frmCart.ITBM.value = DocCur + ' ' + arrData[1];
					document.frmCart.importe.value = DocCur + ' ' + arrData[2];
					loadMinRep();
	   		}
		});
}

function chkListNum(groupNum, Index)
{
	var value = ListNum[Index];
	
	doProcGroupNum(groupNum, value == PList);

	if (value != PList)
	{
		if (confirm("<%=getcartjsLngStr("LtxtApplyPListByPTerm")%>"))
		{
			document.frmCart.NewPList.value = value;
			doApplyPriceList(value);
			PList = value;
		}
	}
}

function goAdd(Confirm, DocConf, Draft, Authorize)
{
	document.frmCart.Confirm.value = Confirm;
	document.frmCart.DocConf.value = DocConf;
	document.frmCart.Draft.value = Draft;
	document.frmCart.Authorize.value = Authorize;
	document.frmCart.I2.click();
}

function goViewItem(lineID, item)
{
	if (UserType == 'C')
	{
		window.location.href='item.asp?Item=' + item + '&cmd=d';
	}
	else
	{
		itemLoadLineID = lineID;
		ItemCmd = 'D';
		openItemDetails(item);
	}
}

function doApplyPriceList(listNum)
{
	var lineNum = '';
	<% If Request("chkLineSumQty") = "true" Then %>
	lineNum = document.frmCart.LineNumDOC1.value;
	<% End If %>
	
	$.post("cart/cartProcess.asp?d=" + (new Date()).toString(), { Field: 'GroupNumAppyList', FieldType: 'N', Value: listNum, LineID: lineNum },
	   function(data){
	     if (data.indexOf('|') == -1)
	     {
	     	alert('<%=getcartjsLngStr("DtxtErrSaveData")%>');
	     }
	     else
	     {
	     	var myArr = data.split('|');
	     	
	     	var arrData = myArr[1].split('{S}');
	     	for (var i = 0;i<arrData.length;i++)
	     	{
	     		var arrLine = arrData[i].split('{C}');
	     		var lineNum = arrLine[0];
	     		var unitPrice = arrLine[1];
	     		var price = arrLine[2];
	     		var currency = arrLine[3];
	     		var discPrcnt = arrLine[4];
	     		var lineTotal = arrLine[5];
	     		
				if (document.getElementById('UnitPrice' + lineNum)) document.getElementById('UnitPrice' + lineNum).value = currency + ' ' + unitPrice;
				if (document.getElementById('price' + lineNum)) document.getElementById('price' + lineNum).value = currency + ' ' + price;
				if (document.getElementById('DiscPrcnt' + lineNum)) document.getElementById('DiscPrcnt' + lineNum).value = '% ' + discPrcnt;
				document.getElementById('LineTotal' + lineNum).value = DocCur + ' ' + lineTotal;
	     	}
	     	
	     	<% If Request("chkLineSumQty") = "true" Then %>
	     	document.getElementById('SummTotal').innerText = DocCur + ' ' + myArr[2];
		    <% End If %>
		     
		     loadTotalVals(myArr[3]);
		     loadMinRep();
	     }
   });
}

function doProc(fld, fldType, value)
{
	$.post("cart/cartProcess.asp?d=" + (new Date()).toString(), { Field: fld, FieldType: fldType, Value: value },
	   function(data){
	   	switch (fld)
	   	{
	   		case 'ShipToCode':
	   		case 'PayToCode':
	   			var arrData = data.split('|');
	   			if (arrData[0] != 'ok')
	   			{
			     	alert('<%=getcartjsLngStr("DtxtErrSaveData")%>');
	   			}
	   			else
	   			{
	   				switch (fld)
	   				{
	   					case 'ShipToCode':
	   						txtShipAddress.innerHTML = arrData[1];
	   						break;
	   					case 'PayToCode':
	   						txtPayAddress.innerHTML = arrData[1];
	   						break;
	   				}
	   			}
	   			break;
	   		case 'DocDueDate':
	   			var arrData = data.split('|');
	   			if (arrData[0] != 'ok')
	   			{
			     	alert('<%=getcartjsLngStr("DtxtErrSaveData")%>');
	   			}
	   			else
	   			{
	   				document.getElementById('DocDueDateAlert').style.display = arrData[1] == 'Y' ? '' : 'none';
	   				if (document.frmCart.I2) document.frmCart.I2.disabled = arrData[2] == 'Y';
	   			}
	   			break;
	   		default:
			     if (data != 'ok')
			     {
			     	alert('<%=getcartjsLngStr("DtxtErrSaveData")%>');
			     }
			     break;
	   	}
   });
}

function GetYesNo(value)
{
	return value ? 'Y' : 'N';
}

function chkKeyDown(t, e, i)
{
	var lineNum = document.frmCart.LineNumDOC1.value.split(', ');
	switch (e.keyCode)
	{
		case 38:
			if (i > 0) 
			{
				try
				{
					if (document.all)
						document.getElementById(getKeyDownID(t) + lineNum[i-1]).focus();
					else
						document.getElementById(getKeyDownID(t) + lineNum[i-1]).focus();
					}
				catch (err)
				{
					alert(err.description + lineNum[i-1]);
				}
			}
			return false;
		case 40:
			if (lineNum.length)
			{
				if (lineNum[i+1])
				{
					try
					{
						if (document.all)
							document.getElementById(getKeyDownID(t) + lineNum[i+1]).focus();
						else
							document.getElementById(getKeyDownID(t) + lineNum[i+1]).focus();
					}
					catch (err)
					{
						alert(err.description + lineNum[i+1]);
					}
				}

			}
			return false;
		case 8:
		case 9:
		case 16:
		case 38:
		case 36:
		case 46:
		case 48:
		case 49:
		case 50:
		case 51:
		case 52:
		case 53:
		case 54:
		case 55:
		case 56:
		case 57:
		case 96:
		case 97:
		case 98:
		case 99:
		case 100:
		case 101:
		case 102:
		case 103:
		case 104:
		case 105:
		case 190: //Punto decimal
		case 110: //Punto decimal
		case 188: //Comma decimal
		case 37: // Left
		case 39: //Right
			return true;
	}
	return false;
}

function getKeyDownID(param)
{
	switch (param)
	{
		case 'Q':
			return 'Qty';
		case 'D':
			return 'DiscPrcnt';
		case 'P':
			return 'price';
	}
}

function doCheckDel(Img, LogNum)
{
	if (!document.getElementById('DelLine' + LogNum).checked)
	{
		document.getElementById('DelLine' + LogNum).checked = true;
		Img.src = 'images/checkbox_on.jpg';
	}
	else
	{
		document.getElementById('DelLine' + LogNum).checked = false;
		Img.src = 'images/checkbox_off.jpg';
	}
}

function setLineQty(LineNum)
{
	var qty = document.getElementById('Qty' + LineNum);
	var DiscPrcnt = document.getElementById('DiscPrcnt' + LineNum);

	$.post("cart/cartProcessLine.asp?d=" + (new Date()).toString(), { LineNum: LineNum, Quantity: qty.value, ProcType: 'Quantity' },
	   function(data){
	   		var myData = data.split('{T}');
	   		var arrData = myData[0].split('{S}');
	   		
	   		var lineQty = arrData[0];
	   		var linePrice = arrData[1];
	   		var lineCur = arrData[2];
	   		var lineDisc = arrData[3];
	   		var lineTotal = arrData[4];
	   		var errOffQty = arrData[5] == 'Y';
	   		var ofertIndex = arrData[6];
	   		var updChilds = arrData[7] == 'Y';
	   		var chkInv = arrData[8] == 'Y';
	   		var lockAdd = arrData[9] == 'Y';
	   		
			if (errOffQty)
			{
				alert('<%=getcartjsLngStr("LtxtOfertQtyErr")%>'.replace('{0}', '<%=Replace(Request("txtOfert"), "'", "\'")%>').replace('{1}', ofertIndex));
			}
	   		
	   		qty.value = lineQty;
	   		if (document.getElementById('price' + LineNum)) document.getElementById('price' + LineNum).value = lineCur + ' ' + linePrice;
	   		if (DiscPrcnt) DiscPrcnt.value = '% ' + lineDisc;
	   		if (document.getElementById('LineTotal' + LineNum)) document.getElementById('LineTotal' + LineNum).value = DocCur + ' ' + lineTotal
	   		
			document.getElementById('InvErr' + LineNum).style.display = chkInv ? 'none' : '';
			if (document.frmCart.I2) document.frmCart.I2.disabled = lockAdd;
			
			if (updChilds)
			{
				var arrChilds = arrData[10].split('{L}');
				for (var i = 1;i<arrChilds.length;i++)
				{
					var arrChild = arrChilds[i].split('{C}');
					var childNum = arrChild[0];
					var childQty = arrChild[1];
					var childTotal = arrChild[2];
					
					document.getElementById('Qty' + childNum).value = childQty;
					if (document.getElementById('LineTotal' + childNum)) document.getElementById('LineTotal' + childNum).value = DocCur + ' ' + childTotal;
					
				}
			}
			
			loadTotalVals(myData[1]);
			
			<% If userType = "V" Then %>loadMinRep();<% End If %>
			
			doReloadLineAddData(LineNum);
	   });
}

function loadTotalVals(data)
{
	var arrData = data.split('{S}');
	
	var docSubTotal = arrData[0];
	var docExpenses = arrData[1];
	var docDisc = arrData[2];
	var docDPM = arrData[3];
	var docTax = arrData[4];
	var docTotal = arrData[5];
	
	document.frmCart.SubTotal.value = DocCur + ' ' + docSubTotal;
	if (document.frmCart.DiscPrcntVal) document.frmCart.DiscPrcntVal.value = DocCur + ' ' + docDisc;
	document.frmCart.ITBM.value = DocCur + ' ' + docTax;
	document.frmCart.importe.value = DocCur + ' ' + docTotal;
	<% If Session("PayCart") Then %>
	var payDocTotal = arrData[5];
	document.frmCart.TotalMC.value = PayDocCur + ' ' + payDocTotal;
	<% End If %>
}

function setLinePrice(LineNum, Currency)
{
	var Price = document.getElementById('price' + LineNum);
	var DiscPrcnt = document.getElementById('DiscPrcnt' + LineNum);
	$.post("cart/cartProcessLine.asp?d=" + (new Date()).toString(), { LineNum: LineNum, Price: Price.value, Cur: Currency, ProcType: 'Price' },
	   function(data){
	   		var myData = data.split('{T}');
	   		var arrData = myData[0].split('{S}');
	   		
	   		var linePrice = arrData[0];
	   		var lineCur = arrData[1];
	   		var lineTotal = arrData[2];
	   		var lineDisc = arrData[3];
			var errMaxDisc = arrData[4] == 'Y';
			
			if (errMaxDisc) alert('<%=getcartjsLngStr("LtxtMaxDiscount")%>'.replace('{0}', lineDisc));
	   		
	   		Price.value = lineCur + ' ' + linePrice;
	   		if (DiscPrcnt) DiscPrcnt.value = '% ' + lineDisc;
	   		document.getElementById('LineTotal' + LineNum).value = DocCur + ' ' + lineTotal
	   		
			loadTotalVals(myData[1]);
			
			<% If userType = "V" Then %>loadMinRep();<% End If %>
			
			doReloadLineAddData(LineNum);
	   });
}

function setLineUn(LineNum)
{
	var selUn = document.getElementById('selUn' + LineNum);
	var qty = document.getElementById('Qty' + LineNum);
	var UnitPrice = document.getElementById('UnitPrice' + LineNum);
	var DiscPrcnt = document.getElementById('DiscPrcnt' + LineNum);

	$.post("cart/cartProcessLine.asp?d=" + (new Date()).toString(), { LineNum: LineNum, Quantity: qty.value, SaleType: selUn.value, ProcType: 'SaleType' },
	   function(data){
	   		var myData = data.split('{T}');
	   		var arrData = myData[0].split('{S}');
	   		
	   		var lineQty = arrData[0];
	   		var linePrice = arrData[1];
	   		var lineUnitPrice = arrData[2];
	   		var lineCur = arrData[3];
	   		var lineDisc = arrData[4];
	   		var lineTotal = arrData[5];
	   		var chkInv = arrData[6] == 'Y';
	   		var lockAdd = arrData[7] == 'Y';
	   		
	   		qty.value = lineQty;
	   		document.getElementById('price' + LineNum).value = lineCur + ' ' + linePrice;
	   		if (UnitPrice) UnitPrice.value = lineCur + ' ' + lineUnitPrice;
	   		if (DiscPrcnt) DiscPrcnt.value = '% ' + lineDisc;
	   		document.getElementById('LineTotal' + LineNum).value = DocCur + ' ' + lineTotal
	   		
			document.getElementById('InvErr' + LineNum).style.display = chkInv ? 'none' : '';
			if (document.frmCart.I2) document.frmCart.I2.disabled = lockAdd;
			
			loadTotalVals(myData[1]);
			
			<% If userType = "V" Then %>loadMinRep();<% End If %>
			
			doReloadLineAddData(LineNum);
	   });
}
function setExpVal(lineNum, fld)
{
	$.post("cart/cartProcess.asp?d=" + (new Date()).toString(), { Field: 'LineTotal', FieldType: 'N', Value: fld.value, LineID: lineNum, LineType: 'E' },
		function(data)
		{
			var arrData = data.split('{S}');
			var expVal = arrData[0];
			var docTax = arrData[1];
			var docTotal = arrData[2];
			
			fld.value = DocCur + ' ' + expVal;
			document.frmCart.ITBM.value = DocCur + ' ' + docTax;
			document.frmCart.importe.value = DocCur + ' ' + docTotal;

			
			<% If userType = "V" Then %>loadMinRep();<% End If %>
		});
}
function setDocDisc()
{
	var DiscPrcnt = document.frmCart.DiscPrcnt;
	
	$.post("cart/cartProcess.asp?d=" + (new Date()).toString(), { Field: 'DocDiscount', FieldType: 'N', Value: DiscPrcnt.value },
		function(data)
		{
	   		var arrData = data.split('{S}');

			var discount = arrData[0];
			var errMaxDisc = arrData[1] == 'Y';
			var docDiscount = arrData[2];
			var docTax = arrData[3];
			var docTotal = arrData[4];
			
			if (errMaxDisc) alert('<%=getcartjsLngStr("LtxtMaxDiscount")%>'.replace('{0}', discount));
			
			DiscPrcnt.value = discount;
			document.frmCart.DiscPrcntVal.value = DocCur + ' ' + docDiscount;
			document.frmCart.ITBM.value = DocCur + ' ' + docTax;
			document.frmCart.importe.value = DocCur + ' ' + docTotal;
			
			<% If userType = "V" Then %>loadMinRep();<% End If %>
		});
}
function setDocDPM()
{
	var DPM = document.frmCart.DpmPrcnt;
	
	$.post("cart/cartProcess.asp?d=" + (new Date()).toString(), { Field: 'DpmPrcnt', FieldType: 'N', Value: DPM.value },
		function(data)
		{
	   		var arrData = data.split('{S}');

			var dpm = arrData[0];
			var errMaxDPM = arrData[1] == 'Y';
			var docDPM = arrData[2];
			var docTax = arrData[3];
			var docTotal = arrData[4];
			
			if (errMaxDPM) alert('<%=getcartjsLngStr("LtxtMaxDPM")%>');
			
			DPM.value = dpm;
			document.frmCart.DPMVal.value = DocCur + ' ' + docDPM;
			document.frmCart.ITBM.value = DocCur + ' ' + docTax;
			document.frmCart.importe.value = DocCur + ' ' + docTotal;
			
			<% If userType = "V" Then %>loadMinRep();<% End If %>
		});
}

function setLineDisc(LineNum)
{
	var DiscPrcnt = document.getElementById('DiscPrcnt' + LineNum);
	$.post("cart/cartProcessLine.asp?d=" + (new Date()).toString(), { LineNum: LineNum, Discount: DiscPrcnt.value, ProcType: 'Discount' },
	   function(data){
	   		var myData = data.split('{T}');
	   		var arrData = myData[0].split('{S}');
	   		
	   		var linePrice = arrData[0];
	   		var lineCur = arrData[1];
	   		var lineTotal = arrData[2];
	   		var lineDisc = arrData[3];
			var errMaxDisc = arrData[4] == 'Y';
			
			if (errMaxDisc) alert('<%=getcartjsLngStr("LtxtMaxDiscount")%>'.replace('{0}', lineDisc));
	   		
	   		document.getElementById('price' + LineNum).value = lineCur + ' ' + linePrice;
	   		DiscPrcnt.value = '% ' + lineDisc;
	   		document.getElementById('LineTotal' + LineNum).value = DocCur + ' ' + lineTotal
	   		
			loadTotalVals(myData[1]);
			
			<% If userType = "V" Then %>loadMinRep();<% End If %>
			
			doReloadLineAddData(LineNum);
	   });
}

function delExp(LineNum, ItemName)
{
	if(confirm('<%=getcartjsLngStr("LtxtConfRemExpense")%>'.replace('{0}', ItemName)))
		window.location.href = 'cart/cartrm.asp?line=' + LineNum + '&exp=Y&redir=cart';
}

function updLineMoreBtn(lineNum, blue)
{
	var addBlue = '';
	if (blue) addBlue = 'blue';
	document.getElementById('btnLineMore' + lineNum).src = 'images/expand' + addBlue + '.gif';
}

function showHeaderDet(btn, hdMark, cartSHAddStr)
{
	for (var i = 0;i<trCartHD.length;i++)
	{
		trCartHD[i].style.display = hdMark.value == 'Y' ? '' : 'none';
	}
	if (hdMark.value == 'Y')
	{
		hdMark.value = 'N';
		if (cartSHAddStr == '')
			btn.value = '- <%=getcartjsLngStr("LtxtHideHdrDet")%>';
		else
			btn.value = cartSHAddStr + '<%=getcartjsLngStr("LtxtHideHdrDet")%>';
		btn.className = 'BtnLess';
	}
	else
	{
		hdMark.value = 'Y';
		if (cartSHAddStr == '')
			btn.value = '+ <%=getcartjsLngStr("LtxtShowHdrDet")%>';
		else
			btn.value = cartSHAddStr + '<%=getcartjsLngStr("LtxtShowHdrDet")%>';
		btn.className = 'BtnMore';
	}
}

function getDType()
{
	if (document.frmCart.R1.length)
	{
		for (var i = 0;i<document.frmCart.R1.length;i++)
		{
			if (document.frmCart.R1[i].checked) return document.frmCart.R1[i].value;
		}
	}
	else
	{
		return document.frmCart.R1.value;
	}
}

function valBtch()
{
	var msgBtch = '<%=getcartjsLngStr("LtxtValIncBtch")%>';
	var msgBtchConf = '\n<%=getcartjsLngStr("LtxtConfContinue")%>';
	var DType = getDType();
	if (DType == '13' || DType == '15' || DType == '17' && VerfyBtchOrder != 'N')
		if (document.frmCart.btnSB)
		{
			if (document.frmCart.btnSB.length)
			{
				for (var i = 0;i<document.frmCart.btnSB.length;i++)
				if (document.frmCart.btnSB[i].src.indexOf('batch_check') == -1)
				{
					if (VerfyBtchOrder == 'C' && DType == '17')
					{
						if (!confirm(msgBtch + msgBtchConf))
						{
							return false;
						}
						else
						{
							return true;
						}
					}
					else
					{
						alert(msgBtch);
						return false;
					}
				}
			}
			else
			{
				if (document.frmCart.btnSB.src.indexOf('batch_check') == -1)
				{
					if (VerfyBtchOrder == 'C' && DType == '17')
					{
						if (!confirm(msgBtch + msgBtchConf))
						{
							return false;
						}
						else
						{
							return true;
						}
					}
					else
					{
						alert(msgBtch);
						return false;
					}
				}
			}
		}
		else if (IncBtch > 0)
		{
			if (VerfyBtchOrder == 'C' && DType == '17')
			{
				if (!confirm(msgBtch + msgBtchConf))
				{
					return false;
				}
				else
				{
					return true;
				}
			}
			else
			{
				alert(msgBtch);
				return false;
			}
		}
	return true;
}

function val3dx()
{
	var msg3dx = '<%=getcartjsLngStr("LtxtValInc3dx")%>';
	var msg3dxConf = '\n<%=getcartjsLngStr("LtxtConfContinue")%>';
	var DType = getDType();
	if (DType == '13' || DType == '15' || DType == '17' && Verfy3dxOrder != 'N')
		if (document.frmCart.btnS3)
		{
			if (document.frmCart.btnS3.length)
			{
				for (var i = 0;i<document.frmCart.btnS3.length;i++)
				if (document.frmCart.btnS3[i].src.indexOf('3dx_check.gif') == -1)
				{
					if (Verfy3dxOrder == 'C' && DType == '17')
					{
						if (!confirm(msg3dx + msg3dxConf))
						{
							return false;
						}
						else
						{
							return true;
						}
					}
					else
					{
						alert(msg3dx);
						return false;
					}
				}
			}
			else
			{
				if (document.frmCart.btnS3.src.indexOf('3dx_check.gif') == -1)
				{
					if (Verfy3dxOrder == 'C' && DType == '17')
					{
						if (!confirm(msg3dx + msg3dxConf))
						{
							return false;
						}
						else
						{
							return true;
						}
					}
					else
					{
						alert(msg3dx);
						return false;
					}
				}
			}
		}
		else if (Inc3dx > 0)
		{
			if (Verfy3dxOrder == 'C' && DType == '17')
			{
				if (!confirm(msg3dx + msg3dxConf))
				{
					return false;
				}
				else
				{
					return true;
				}
			}
			else
			{
				alert(msg3dx);
				return false;
			}
		}
	return true;
}

//popup para campo SDK de imagen
function getImg(Field, Img, MaxSize)
{
	SaveImgField = Field;
	SaveImgImage = Img;
	SaveImgMaxSize = MaxSize;
	OpenWin = this.open('upload/fileupload.aspx?ID=<%=Session("ID")%>&style=../design/<%=Request("SelDes")%>/style/stylePopUp.css', "ImagePicker", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no, width=300,height=111");
}

//funcion para devolver nombre de imagen de popup
function changepic(img_src) {
SaveImgField.value = img_src;
SaveImgImage.src = "pic.aspx?filename=" + img_src + "&MaxSize=" + SaveImgMaxSize+'&dbName=<%=Session("olkdb")%>';
doProc(SaveImgField.name, 'S', img_src);
}

function doSB(img, Item, sbType, LineNum)
{
	SaveImgImage = img;
	page = 'cart/setCart' + sbType + '.asp?LineNum=' + LineNum + "&Item=" + Item + '&pop=Y&AddPath=';
	OpenWin = this.open(page, "OpenWin", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no, width=600,height=400");
	OpenWin.focus()
}

function setSBImg(cmd)
{
	switch (cmd)
	{
		case 0:
			SaveImgImage.alt = '<%=getcartjsLngStr("LtxtIncompleteBatchs")%>';
			SaveImgImage.src = 'images/batch_nocheck.gif';
			break;
		case 1:
			SaveImgImage.alt = '<%=getcartjsLngStr("LtxtViewBtchSum")%>';
			SaveImgImage.src = 'images/batch_check.gif';
			break;
		case 2:
			SaveImgImage.alt = '<%=getcartjsLngStr("LtxtIncompleteBatchs")%>';
			SaveImgImage.src = 'images/batch_checkGris.gif';
			break;
	}
}

function setS3Img(cmd)
{
	switch (cmd)
	{
		case 0:
			SaveImgImage.alt = '<%=getcartjsLngStr("LtxtIncomplete3dx")%>';
			SaveImgImage.src = 'images/3dx_nocheck.gif';
			break;
		case 1:
			SaveImgImage.alt = '<%=getcartjsLngStr("LtxtView3dxSum")%>';
			SaveImgImage.src = 'images/3dx_check.gif';
			break;
		case 2:
			SaveImgImage.alt = '<%=getcartjsLngStr("LtxtIncomplete3dx")%>';
			SaveImgImage.src = 'images/3dx_checkGris.gif';
			break;
	}
}

function setSSImg(cmd)
{
	switch (cmd)
	{
		case 0:
			SaveImgImage.alt = '<%=getcartjsLngStr("LtxtIncompleteSeries")%>';
			SaveImgImage.src = 'images/serial_nocheck.gif';
			break;
		case 1:
			SaveImgImage.alt = '<%=getcartjsLngStr("LtxtViewSeries")%>';
			SaveImgImage.src = 'images/serial_check.gif';
			break;
		case 2:
			SaveImgImage.alt = '<%=getcartjsLngStr("LtxtIncompleteSeries")%>';
			SaveImgImage.src = 'images/serial_checkGris.gif';
			break;
	}
}

//Variables para cambio de tipo de documento
var DueDate13 = '<%=getcartjsLngStr("LtxtPymntDue")%>';
var DueDate15 = '<%=getcartjsLngStr("LtxtDelDate")%>';
var DueDate17 = '<%=getcartjsLngStr("LtxtDelDate")%>';
var DueDate23 = '<%=getcartjsLngStr("LtxtComDate")%>';
var Confirm13 = '<% If myAut.GetObjectProperty(13, "C") Then %><%=getcartjsLngStr("DtxtConfirm")%><% Else %><%=getcartjsLngStr("DtxtAdd")%><% End If %>';
var Confirm_13 = '<% If myAut.GetObjectProperty(-13, "C") Then %><%=getcartjsLngStr("DtxtConfirm")%><% Else %><%=getcartjsLngStr("DtxtAdd")%><% End If %>';
var Confirm15 = '<% If myAut.GetObjectProperty(15, "C") Then %><%=getcartjsLngStr("DtxtConfirm")%><% Else %><%=getcartjsLngStr("DtxtAdd")%><% End If %>';
var Confirm17 = '<% If myAut.GetObjectProperty(17, "C") Then %><%=getcartjsLngStr("DtxtConfirm")%><% Else %><%=getcartjsLngStr("DtxtAdd")%><% End If %>';
var Confirm23 = '<% If myAut.GetObjectProperty(23, "C") Then %><%=getcartjsLngStr("DtxtConfirm")%><% Else %><%=getcartjsLngStr("DtxtAdd")%><% End If %>';

//Variables de moneda
var DocCur;
var MainCur;
var PayDocCur = "<%=Request("PayDocCur")%>";
var oldDocCur;

function changeDocDate()
{
	var DocDate = document.frmCart.DocDate;
	
	$.post("cart/cartProcess.asp?d=" + (new Date()).toString(), { Field: 'DocDate', FieldType: 'D', Value: DocDate.value },
		function(data)
		{
	   		var arrData = data.split('{S}');

			var docDueDate = arrData[0];
			var errDocDate = arrData[1] == 'Y';
			var docTax = arrData[2];
			var docTotal = arrData[3];
			
			document.frmCart.DocDueDate.value = docDueDate;
			document.getElementById('DocDateAlert').style.display = errDocDate ? '' : 'none';
			document.getElementById('DocDueDateAlert').style.display = 'none';
			
			document.frmCart.ITBM.value = DocCur + ' ' + docTax;
			document.frmCart.importe.value = DocCur + ' ' + docTotal;
			
			CheckIsCartLocked();
			
			<% If userType = "V" Then %>loadMinRep();<% End If %>
		});
}

function changeDocDueDate()
{
	doProc('DocDueDate', 'D', document.frmCart.DocDueDate.value);
}

//Funcion al cambiar tipo de documento cambiar texto a boton de agregar/confirmar y texto de cambio txtDueDate
function changeDocType(Type)
{
	$.post("cart/cartProcess.asp?d=" + (new Date()).toString(), { Field: 'ObjectCode', FieldType: 'N', Value: Type },
		function(data)
		{
			var arrData = data.split('{S}');

			var docDueDate = arrData[0];
			var chkInvOp = arrData[1];
			
			document.frmCart.DocDueDate.value = docDueDate;
			document.getElementById('DocDueDateAlert').style.display = 'none';
			cartObject = parseInt(Type);
			
			if (document.getElementById('trPartSupply')) document.getElementById('trPartSupply').style.display = cartObject == 17 ? '' : 'none';
			
			if (chkInvOp != 'N') doCheckLinesInv(chkInvOp);
			
			<% If userType = "V" Then %>loadMinRep();<% End If %>
		});

	var txtDocDueDateDesc;
	var txtBtnAdd;
	var showSetLnk = '';
	switch (parseInt(Type))
	{
		case -13:
			txtDocDueDateDesc = DueDate13;
			txtBtnAdd = Confirm_13;
			break;
		case 15:
			txtDocDueDateDesc = DueDate15;
			txtBtnAdd = Confirm15;
			break;
		case 17:
			txtDocDueDateDesc = DueDate17;
			txtBtnAdd = Confirm17;
			break;
		case 23:
			txtDocDueDateDesc = DueDate23;
			txtBtnAdd = Confirm23;
			showSetLnk = 'none';
			break;
		default:
			txtDocDueDateDesc = DueDate13;
			txtBtnAdd = Confirm13;
			break;
	}
	txtDocDueDateDesc += ':';
	document.getElementById('txtDocDueDate').innerHTML = '<nobr><b><font size="1" face="Verdana">' + txtDocDueDateDesc + '<font color="red">*</font></font></b></nobr>';
	document.getElementById('DocDueDateAlert').setAttribute('alt', txtDocDueDateLimit.replace('{0}', txtDocDueDateDesc.toLowerCase()));
	
	<% If Request("Verfy") = "True" and userType = "V" Then %>if (document.frmCart.I2) document.frmCart.I2.value = txtBtnAdd;<% End If %>
	if (document.getElementById('tdSetLnk'))
	{
		if (document.getElementById('tdSetLnk').length)
		{
			for (var i = 0;i<document.getElementById('tdSetLnk').length;i++)
			{
				document.getElementById('tdSetLnk')[i].style.display = showSetLnk;
			}
		}
		else if (document.getElementById('tdSetLnk'))
		{
			document.getElementById('tdSetLnk').style.display = showSetLnk;
		}
	}
}

function doCheckLinesInv(op)
{
	var LineNumDOC1 = document.frmCart.LineNumDOC1.value;
	switch (op)
	{
		case 'C':
			if (LineNumDOC1 != '')
			{
				var lineNum = LineNumDOC1.split(', ');
				for (var i = 0;i<lineNum.length;i++)
				{
					document.getElementById('InvErr' + lineNum[i]).style.display = 'none';
				}
			}
			<% If Request("ViewMode") <> "all" and Request("chkLineSumQty") = "true" Then %>
			var iconAlertLineSum = document.getElementById('iconAlertLineSum');
			if (iconAlertLineSum) iconAlertLineSum.style.display = 'none';
			<% End If %>
			break;
		case 'Y':
			$.post('cart/cartProcess.asp', { Field: 'CheckLinesInv', Lines: LineNumDOC1 }, function doLinesInv(data) 
			{ 
				if (data != '')
				{
					var arrData = data.split('{S}');
					for (var i = 0;i<arrData.length;i++)
					{
						var lineData = arrData[i].split('{C}');
						var lineNum = lineData[0];
						var chkInv = lineData[1];
						
						document.getElementById('InvErr' + lineNum).style.display = parseInt(chkInv) == 0 ? 'none' : '';
					}
				}
			});
			<% If Request("ViewMode") <> "all" and Request("chkLineSumQty") = "true" Then %>
			$.post('cart/cartProcess.asp', { Field: 'CheckSumInv', Lines: LineNumDOC1 }, function doSumLineInv(data) 
			{ 
				var iconAlertLineSum = document.all ? document.getElementById('iconAlertLineSum') : document.getElementById('iconAlertLineSum');
				if (parseInt(data) > 0)
				{
					iconAlertLineSum.style.display = '';
					iconAlertLineSum.alt = txtErrSumItmInv.replace('{0}', data);
				}
				else
				{
					iconAlertLineSum.style.display = 'none';
				}
			});
			<% End If %>
			break;
	}
	CheckIsCartLocked();
}

function CheckIsCartLocked()
{
	$.post('cart/cartProcess.asp', { Field: 'IsLocked' }, function doCheckLock(data) 
	{ 
		if (document.frmCart.I2) document.frmCart.I2.disabled = data == 'Y';
	});
}

//Variables de cliente
var Balance = getNumeric('<%=Request("Balance")%>');
var CreditLimit = getNumeric('<%=Request("CreditLine")%>');
var finalTotal = 0;

//Campo temporal para darle valor luego que un popup devolvio un valor
var objField;

//Funcion de popup general
function Start(page) {
<% If userType = "C" Then %>
OpenWin = this.open(page, "ItemDetails", "toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no,width=482,height=450");
<% ElseIf userType = "V" Then %>
OpenWin = this.open(page, "ItemDetails", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=no, width=598,height=420");
<% End If %>
OpenWin.focus()
}

//Segunda funcion de popup con  solo variable de altura
function Start2(page, h, s) {
OpenWin = this.open(page, "LineNote", "toolbar=no,menubar=no,location=no,scrollbars=" + s + ",resizable=no, width=400,height=" + h);
OpenWin.focus()
}
var objFieldType;
//Funcion de popup para llamar a datepicker
function datePicker(page, w, h, s, r, o, procType) {
objField = o;
objFieldType = procType;
OpenWin = this.open(page, "datePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable="+r+", width="+w+",height="+h);
OpenWin.focus();
}

//Funcion para ponele el valor seleccionad del datepicker al objField
function setTimeStamp(Action, varDate) { objField.value = varDate; doProc(objField.name, objFieldType, varDate); }

<% If Session("PayCart") then %>
//funcion para realizar pagos en canasta con recibo
function Pay(theURL, popW, popH, scroll) { // V 1.0
var winleft = (screen.width - popW) / 2;
var winUp = (screen.height - popH) / 2;
winProp = 'width='+popW+',height='+popH+',left='+winleft+',top='+winUp+',toolbar=no,scrollbars='+scroll+',menubar=no,location=no,resizable=no'
theURL2 = theURL+'?imp='+document.frmCart.TotalMC.value+'&pag='+document.frmCart.pagado.value+'&voucher=NULL&saldofuera=False&pop=Y&AddPath=&EnableMC=<%=Request("EnableMC")%>'
OpenWin = window.open(theURL2, "cashCart", winProp)
}

//Cuando cierro un metodo de pago actualizo la cantidad pagada en la canasta
function updatePagado(pagVal)
{ document.frmCart.pagado.value = PayDocCur + ' ' + pagVal }
<% end if %>

//Funcion de calculo de totalos al cambiar moneda en documento
function changeCur() {
	var lineNum = '';
	<% If Request("chkLineSumQty") = "true" Then %>
	lineNum = document.frmCart.LineNumDOC1.value;
	<% End If %>
	
	var selCur = document.frmCart.DocCur.value;

	$.post("cart/cartProcess.asp?d=" + (new Date()).toString(), { Field: 'DocCur', FieldType: 'S', Value: selCur, LineID: lineNum },
		function(data)
		{
			var arrData = data.split('{S}');
			DocCur = document.frmCart.DocCur.value;	
			
			if (arrData[0] != '')
			{
				var arrLinesData = arrData[0].split('{L}');
				for (var i = 0;i<arrLinesData.length;i++)
				{
					var lineData = arrLinesData[i].split('{C}');
					var lineNum = lineData[0];
					var lineTotal = lineData[1];
					
					document.getElementById('LineTotal' + lineNum).value = DocCur + ' ' + lineTotal;
				}
			}
			
			<% If Request("chkLineSumQty") = "true" Then %>
			document.getElementById('SummTotal').innerText = DocCur + ' ' + arrData[1];
			<% End If %>
			
			if (arrData[2] != '')
			{
				var arrExp = arrData[2].split('{L}');
				for (var i = 0;i<arrExp.length;i++)
				{
					var lineData = arrExp[i].split('{C}');
					var lineNum = lineData[0];
					var lineTotal = lineData[1];
					
					document.getElementById('<% If myApp.SVer >= "6" Then Response.Write "Exp" %>Price' + lineNum).value = DocCur + ' ' + lineTotal;
				}
			}
			
			var subTotal = arrData[3];
			var docDiscount = arrData[4];
			var docTax = arrData[5];
			var docTotal = arrData[6];
			var curRate = arrData[8];
			
			document.frmCart.SubTotal.value = DocCur + ' ' + subTotal;
			document.frmCart.DiscPrcntVal.value = DocCur + ' ' + docDiscount;
			document.frmCart.ITBM.value = DocCur + ' ' + docTax;
			document.frmCart.importe.value = DocCur + ' ' + docTotal;
			
			var DocCurRate = document.getElementById('DocCurRate');
			DocCurRate.style.display = MainCur != selCur ? '' : 'none';
			DocCurRate.value = curRate;

			<% If userType = "V" Then %>loadMinRep();<% End If %>
		});
}

//Validacion de eliminación de articulos de canasta
function valDel()
{
	<% If Request("Verfy") Then %>
	dCount = 0;
	if (document.frmCart.DelLine.length)
	{
		for (var i = 0;i<document.frmCart.DelLine.length;i++)
		{
			if (document.frmCart.DelLine(i).checked) dCount++;
		}
	}
	else
	{
		if (document.frmCart.DelLine.checked) dCount++;
	}
	if (dCount > 0)
	{
		if (!confirm('<%=getcartjsLngStr("LtxtConfDelItm")%>'.replace('{0}', dCount)))
		{
			return false;
		}
	}
	if (dCount == 0) { alert('<%=getcartjsLngStr("LtxtDelItmSel")%>'); return false; }
	<% End If %>
}

//Validar canasta antes de agregar
function chkCart(I)
{
	<% If Request("CreditLimit") and (Request("object") = 17 and myApp.OrderLimit or Request("object") = 13 and myApp.SalesLimit or Request("object") = 15 and myApp.DlnLimit) Then %>
	if ((parseFloat(finalTotal)+parseFloat(Balance)) > parseFloat(CreditLimit)) 
	{
		<% 	If userType = "V" Then
				If Session("useraccess") = "U" Then %>
				disableChkWin = true;
				var credDif = DocCur + OLKFormatNumber(parseFloat(finalTotal)+parseFloat(Balance)-parseFloat(CreditLimit),SumDec);
				if (!confirm('<%=getcartjsLngStr("LtxtValCreditLimit")%>'.replace('{0}', '<%=myHTMLEncode(Request("txtClient"))%>').replace('{1}', credDif) + '\n' + '<%=getcartjsLngStr("LtxtConfContinue")%>'))
				{
					disableChkWin = false;
					return false;
				}
				disableChkWin = false;
				<% Else %>
				disableChkWin = true;
				var credDif = DocCur + OLKFormatNumber(parseFloat(finalTotal)+parseFloat(Balance)-parseFloat(CreditLimit),SumDec);
				if (!confirm('<%=getcartjsLngStr("LtxtExceedClientCredi")%>'.replace('{0}', '<%=myHTMLEncode(Request("txtClient"))%>').replace('{1}', credDif)+ '\n' + '<%=getcartjsLngStr("LtxtConfContinue")%>'))
				{
					disableChkWin = false;
					return false;
				}
				disableChkWin = false;
				<% End If
		ElseIf userType = "C" Then %>
				disableChkWin = true;
				var credDif = DocCur + OLKFormatNumber(parseFloat(finalTotal)+parseFloat(Balance)-parseFloat(CreditLimit),SumDec);
				alert('<%=getcartjsLngStr("LtxtClientExceedStop")%>'.replace('{0}', credDif));
				disableChkWin = false;
				return false;
	<%	End If %>
	}
	<% End If %>
	<% If Session("PayCart") Then %>
	disableChkWin = true;
	if (parseFloat(getNumeric(document.frmCart.pagado.value.replace(PayDocCur, ''))) <= 0)
	{
		alert('<%=getcartjsLngStr("LtxtValPymntElements")%>');
		disableChkWin = false;
		return false;
	}
	else if (parseFloat(getNumeric(document.frmCart.pagado.value.replace(PayDocCur, ''))) > parseFloat(getNumeric(document.frmCart.TotalMC.value.replace(PayDocCur, ''))))
	{
		varDif = OLKFormatNumber(parseFloat(getNumeric(document.frmCart.pagado.value.replace(PayDocCur, '')))-parseFloat(getNumeric(document.frmCart.TotalMC.value.replace(PayDocCur, ''))), SumDec);
		if (!confirm('<%=getcartjsLngStr("LtxtImpRctMoreThenInv")%>'.replace('{0}', '<%=myHTMLEncode(Request("txtInv"))%>').replace('{1}', DocCur + varDif) + '\n<%=getcartjsLngStr("LtxtConfContinue")%>'))
		{
			disableChkWin = false;
			return false;
		}
	}
	disableChkWin = false;
	<% End If %>
	document.frmCart.cartSubmit.value = I;
	setFlowAlertVars('D3', '', 'document.frmCart.submit();', 'cartSubmit.asp?RetVal=<%=Session("RetVal")%>&Confirm=Y&Flow=Y');
	flowDraftFld = 'document.frmCart.Draft.value';
	flowAutFld = 'document.frmCart.Authorize.value';
	doFlowAlert();
}

//Agregar item desde gastos o mejores 10 ventas
function goAddSmallList(value)
{
	if (value != '')
	{
		doMyLink('cart/addCartSubmitExp.asp', 'Item='+value+'&redir=<%=Request("cmd")%>&AddPath=', '');
  	}
}  

//Cambio de lista en canasta
function changeSmallList(Group)
{
	for (var i = document.frmCart.addItemCode.length-1;i>=0;i--)
	{
		document.frmCart.addItemCode.remove(i);
	}
	switch(Group)
	{
		case "G":
			document.getElementById('MinArtTitle').innerHTML = '<%=getcartjsLngStr("LtxtExpenses")%>';
			document.getElementById('iSetData').src = "cart/cartGetSmallList.asp?Group=G";
			break;
		case "V":
			document.getElementById('MinArtTitle').innerHTML = "<%=Replace(getcartjsLngStr("LtxtXItemsSelld"), "{0}", myApp.Top10Items)%>";
			document.getElementById('iSetData').src = "cart/cartGetSmallList.asp?Group=V";
			break;
	}
}
function getAddItemCode() { return document.frmCart.addItemCode; }
function showVolRep(img, Item, e, LineNum)
{
	if (document.getElementById('tblCartVolRep'))
	{
		tblVolRep = document.getElementById('tblCartVolRep');
	}
	else
	{
		CreateCartVolRep(true, 'tblCartVolRep');
		tblVolRep = document.getElementById('tblCartVolRep');
	}

	var volUnit = parseInt(document.getElementById('selUn' + LineNum).value);
	var docDate = document.frmCart.DocDate.value;
	tblVolRepAddTop = 0;
	tblVolRepAddLeft = 0;
	displayVolRep(Item, img, e, volUnit, docDate);
}

function doLineLink(id, line)
{
	$.post('cartGetLinkDataFetch.asp?d=' + (new Date()).toString(), { ID: id, LineNum: line }, function(data)
	{
		var arr = data.split('{R}');
		
		var frm = document.getElementById('frmRSLink');
		if (frm.hasChildNodes())
		{
		    while (frm.childNodes.length >= 1)
		    {
		        frm.removeChild(frm.firstChild);       
		    } 
		}
		
		for (var i = 0;i<arr.length - 1;i++)
		{
			var vals = arr[i].split('{C}');
			
			var obj = document.createElement('input');
			obj.setAttribute('type', 'hidden');
			obj.setAttribute('name', vals[0]);
			obj.setAttribute('value', vals[1]);
			
			frm.appendChild(obj);
			
		}
		document.frmRSLink.submit();
	});
}

function doReloadLineAddData(line)
{
	if (hasDynamicCols)
	{
		$.post('cartGetAddDataFetch.asp?d=' + (new Date()).toString(), { LineNum: line }, function(data)
		{
			var arr = data.split('{R}');
			for (var i = 1;i<arr.length-1;i++)
			{
				var vals = arr[i].split('{C}');
				
				document.getElementById('ci' + line + '_' + vals[0].replace('col', '')).innerHTML = vals[1];
			}
		});
	}
}
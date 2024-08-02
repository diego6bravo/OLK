<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cuentas.xml"
set doccuentas = server.CreateObject("MSXML2.DOMDocument")
doccuentas.async = False
Doccuentas.Load(server.MapPath(xmlfilename)) 
doccuentas.setProperty "SelectionLanguage", "XPath"
set selectedcuentasnode = doccuentas.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcuentasnodes=doccuentas.documentElement.selectNodes("/languages/language")
function getcuentasLngStr(instring)
	temp = selectedcuentasnode.selectSingleNode(instring).text
	getcuentasLngStr = temp
end function
%>

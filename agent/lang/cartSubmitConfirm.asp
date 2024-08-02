<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartSubmitConfirm.xml"
set doccartSubmitConfirm = server.CreateObject("MSXML2.DOMDocument")
doccartSubmitConfirm.async = False
DoccartSubmitConfirm.Load(server.MapPath(xmlfilename)) 
doccartSubmitConfirm.setProperty "SelectionLanguage", "XPath"
set selectedcartSubmitConfirmnode = doccartSubmitConfirm.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartSubmitConfirmnodes=doccartSubmitConfirm.documentElement.selectNodes("/languages/language")
function getcartSubmitConfirmLngStr(instring)
	temp = selectedcartSubmitConfirmnode.selectSingleNode(instring).text
	getcartSubmitConfirmLngStr = temp
end function
%>

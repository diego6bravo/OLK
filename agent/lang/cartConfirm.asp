<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartConfirm.xml"
set doccartConfirm = server.CreateObject("MSXML2.DOMDocument")
doccartConfirm.async = False
DoccartConfirm.Load(server.MapPath(xmlfilename)) 
doccartConfirm.setProperty "SelectionLanguage", "XPath"
set selectedcartConfirmnode = doccartConfirm.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartConfirmnodes=doccartConfirm.documentElement.selectNodes("/languages/language")
function getcartConfirmLngStr(instring)
	temp = selectedcartConfirmnode.selectSingleNode(instring).text
	getcartConfirmLngStr = temp
end function
%>

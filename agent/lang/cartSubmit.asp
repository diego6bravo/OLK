<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartSubmit.xml"
set doccartSubmit = server.CreateObject("MSXML2.DOMDocument")
doccartSubmit.async = False
DoccartSubmit.Load(server.MapPath(xmlfilename)) 
doccartSubmit.setProperty "SelectionLanguage", "XPath"
set selectedcartSubmitnode = doccartSubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartSubmitnodes=doccartSubmit.documentElement.selectNodes("/languages/language")
function getcartSubmitLngStr(instring)
	temp = selectedcartSubmitnode.selectSingleNode(instring).text
	getcartSubmitLngStr = temp
end function
%>

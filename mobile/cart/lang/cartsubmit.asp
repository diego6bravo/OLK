<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartsubmit.xml"
set doccartsubmit = server.CreateObject("MSXML2.DOMDocument")
doccartsubmit.async = False
Doccartsubmit.Load(server.MapPath(xmlfilename)) 
doccartsubmit.setProperty "SelectionLanguage", "XPath"
set selectedcartsubmitnode = doccartsubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartsubmitnodes=doccartsubmit.documentElement.selectNodes("/languages/language")
function getcartsubmitLngStr(instring)
	temp = selectedcartsubmitnode.selectSingleNode(instring).text
	getcartsubmitLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "crcerror.xml"
set doccrcerror = server.CreateObject("MSXML2.DOMDocument")
doccrcerror.async = False
Doccrcerror.Load(server.MapPath(xmlfilename)) 
doccrcerror.setProperty "SelectionLanguage", "XPath"
set selectedcrcerrornode = doccrcerror.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcrcerrornodes=doccrcerror.documentElement.selectNodes("/languages/language")
function getcrcerrorLngStr(instring)
	temp = selectedcrcerrornode.selectSingleNode(instring).text
	getcrcerrorLngStr = temp
end function
%>

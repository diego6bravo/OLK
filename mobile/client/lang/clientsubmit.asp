<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "clientsubmit.xml"
set docclientsubmit = server.CreateObject("MSXML2.DOMDocument")
docclientsubmit.async = False
Docclientsubmit.Load(server.MapPath(xmlfilename)) 
docclientsubmit.setProperty "SelectionLanguage", "XPath"
set selectedclientsubmitnode = docclientsubmit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedclientsubmitnodes=docclientsubmit.documentElement.selectNodes("/languages/language")
function getclientsubmitLngStr(instring)
	temp = selectedclientsubmitnode.selectSingleNode(instring).text
	getclientsubmitLngStr = temp
end function
%>

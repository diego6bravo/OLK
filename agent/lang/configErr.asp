<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "configErr.xml"
set docconfigErr = server.CreateObject("MSXML2.DOMDocument")
docconfigErr.async = False
DocconfigErr.Load(server.MapPath(xmlfilename)) 
docconfigErr.setProperty "SelectionLanguage", "XPath"
set selectedconfigErrnode = docconfigErr.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedconfigErrnodes=docconfigErr.documentElement.selectNodes("/languages/language")
function getconfigErrLngStr(instring)
	temp = selectedconfigErrnode.selectSingleNode(instring).text
	getconfigErrLngStr = temp
end function
%>

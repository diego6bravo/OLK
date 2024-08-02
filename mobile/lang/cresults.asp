<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cresults.xml"
set doccresults = server.CreateObject("MSXML2.DOMDocument")
doccresults.async = False
Doccresults.Load(server.MapPath(xmlfilename)) 
doccresults.setProperty "SelectionLanguage", "XPath"
set selectedcresultsnode = doccresults.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcresultsnodes=doccresults.documentElement.selectNodes("/languages/language")
function getcresultsLngStr(instring)
	temp = selectedcresultsnode.selectSingleNode(instring).text
	getcresultsLngStr = temp
end function
%>

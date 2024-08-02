<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cxcData.xml"
set doccxcData = server.CreateObject("MSXML2.DOMDocument")
doccxcData.async = False
DoccxcData.Load(server.MapPath(xmlfilename)) 
doccxcData.setProperty "SelectionLanguage", "XPath"
set selectedcxcDatanode = doccxcData.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcxcDatanodes=doccxcData.documentElement.selectNodes("/languages/language")
function getcxcDataLngStr(instring)
	temp = selectedcxcDatanode.selectSingleNode(instring).text
	getcxcDataLngStr = temp
end function
%>

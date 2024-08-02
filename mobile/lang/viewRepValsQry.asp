<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "viewRepValsQry.xml"
set docviewRepValsQry = server.CreateObject("MSXML2.DOMDocument")
docviewRepValsQry.async = False
DocviewRepValsQry.Load(server.MapPath(xmlfilename)) 
docviewRepValsQry.setProperty "SelectionLanguage", "XPath"
set selectedviewRepValsQrynode = docviewRepValsQry.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedviewRepValsQrynodes=docviewRepValsQry.documentElement.selectNodes("/languages/language")
function getviewRepValsQryLngStr(instring)
	temp = selectedviewRepValsQrynode.selectSingleNode(instring).text
	getviewRepValsQryLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "executeConf.xml"
set docexecuteConf = server.CreateObject("MSXML2.DOMDocument")
docexecuteConf.async = False
DocexecuteConf.Load(server.MapPath(xmlfilename)) 
docexecuteConf.setProperty "SelectionLanguage", "XPath"
set selectedexecuteConfnode = docexecuteConf.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedexecuteConfnodes=docexecuteConf.documentElement.selectNodes("/languages/language")
function getexecuteConfLngStr(instring)
	temp = selectedexecuteConfnode.selectSingleNode(instring).text
	getexecuteConfLngStr = temp
end function
%>

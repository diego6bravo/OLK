<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "agentTop.xml"
set docagentTop = server.CreateObject("MSXML2.DOMDocument")
docagentTop.async = False
DocagentTop.Load(server.MapPath(xmlfilename)) 
docagentTop.setProperty "SelectionLanguage", "XPath"
set selectedagentTopnode = docagentTop.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedagentTopnodes=docagentTop.documentElement.selectNodes("/languages/language")
function getagentTopLngStr(instring)
	temp = selectedagentTopnode.selectSingleNode(instring).text
	getagentTopLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "ofertAgentContraOfert.xml"
set docofertAgentContraOfert = server.CreateObject("MSXML2.DOMDocument")
docofertAgentContraOfert.async = False
DocofertAgentContraOfert.Load(server.MapPath(xmlfilename)) 
docofertAgentContraOfert.setProperty "SelectionLanguage", "XPath"
set selectedofertAgentContraOfertnode = docofertAgentContraOfert.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedofertAgentContraOfertnodes=docofertAgentContraOfert.documentElement.selectNodes("/languages/language")
function getofertAgentContraOfertLngStr(instring)
	temp = selectedofertAgentContraOfertnode.selectSingleNode(instring).text
	getofertAgentContraOfertLngStr = temp
end function
%>

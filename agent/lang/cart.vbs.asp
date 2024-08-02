<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cart.vbs.xml"
set doccartvbs = server.CreateObject("MSXML2.DOMDocument")
doccartvbs.async = False
Doccartvbs.Load(server.MapPath(xmlfilename)) 
doccartvbs.setProperty "SelectionLanguage", "XPath"
set selectedcartvbsnode = doccartvbs.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartvbsnodes=doccartvbs.documentElement.selectNodes("/languages/language")
function getcartvbsLngStr(instring)
	temp = selectedcartvbsnode.selectSingleNode(instring).text
	getcartvbsLngStr = temp
end function
%>

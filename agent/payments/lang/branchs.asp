<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "branchs.xml"
set docbranchs = server.CreateObject("MSXML2.DOMDocument")
docbranchs.async = False
Docbranchs.Load(server.MapPath(xmlfilename)) 
docbranchs.setProperty "SelectionLanguage", "XPath"
set selectedbranchsnode = docbranchs.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedbranchsnodes=docbranchs.documentElement.selectNodes("/languages/language")
function getbranchsLngStr(instring)
	temp = selectedbranchsnode.selectSingleNode(instring).text
	getbranchsLngStr = temp
end function
%>

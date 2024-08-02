<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartrm.xml"
set doccartrm = server.CreateObject("MSXML2.DOMDocument")
doccartrm.async = False
Doccartrm.Load(server.MapPath(xmlfilename)) 
doccartrm.setProperty "SelectionLanguage", "XPath"
set selectedcartrmnode = doccartrm.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartrmnodes=doccartrm.documentElement.selectNodes("/languages/language")
function getcartrmLngStr(instring)
	temp = selectedcartrmnode.selectSingleNode(instring).text
	getcartrmLngStr = temp
end function
%>

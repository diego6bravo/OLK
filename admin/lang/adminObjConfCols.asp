<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminObjConfCols.xml"
set docadminObjConfCols = server.CreateObject("MSXML2.DOMDocument")
docadminObjConfCols.async = False
DocadminObjConfCols.Load(server.MapPath(xmlfilename)) 
docadminObjConfCols.setProperty "SelectionLanguage", "XPath"
set selectedadminObjConfColsnode = docadminObjConfCols.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminObjConfColsnodes=docadminObjConfCols.documentElement.selectNodes("/languages/language")
function getadminObjConfColsLngStr(instring)
	temp = selectedadminObjConfColsnode.selectSingleNode(instring).text
	getadminObjConfColsLngStr = temp
end function
%>

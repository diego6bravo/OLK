<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "nodata.xml"
set docnodata = server.CreateObject("MSXML2.DOMDocument")
docnodata.async = False
Docnodata.Load(server.MapPath(xmlfilename)) 
docnodata.setProperty "SelectionLanguage", "XPath"
set selectednodatanode = docnodata.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectednodatanodes=docnodata.documentElement.selectNodes("/languages/language")
function getnodataLngStr(instring)
	temp = selectednodatanode.selectSingleNode(instring).text
	getnodataLngStr = temp
end function
%>

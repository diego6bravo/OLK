<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "repLegend.xml"
set docrepLegend = server.CreateObject("MSXML2.DOMDocument")
docrepLegend.async = False
DocrepLegend.Load(server.MapPath(xmlfilename)) 
docrepLegend.setProperty "SelectionLanguage", "XPath"
set selectedrepLegendnode = docrepLegend.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedrepLegendnodes=docrepLegend.documentElement.selectNodes("/languages/language")
function getrepLegendLngStr(instring)
	temp = selectedrepLegendnode.selectSingleNode(instring).text
	getrepLegendLngStr = temp
end function
%>

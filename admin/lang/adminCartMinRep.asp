<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCartMinRep.xml"
set docadminCartMinRep = server.CreateObject("MSXML2.DOMDocument")
docadminCartMinRep.async = False
DocadminCartMinRep.Load(server.MapPath(xmlfilename)) 
docadminCartMinRep.setProperty "SelectionLanguage", "XPath"
set selectedadminCartMinRepnode = docadminCartMinRep.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCartMinRepnodes=docadminCartMinRep.documentElement.selectNodes("/languages/language")
function getadminCartMinRepLngStr(instring)
	temp = selectedadminCartMinRepnode.selectSingleNode(instring).text
	getadminCartMinRepLngStr = temp
end function
%>

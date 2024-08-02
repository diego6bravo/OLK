<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCatProp.xml"
set docadminCatProp = server.CreateObject("MSXML2.DOMDocument")
docadminCatProp.async = False
DocadminCatProp.Load(server.MapPath(xmlfilename)) 
docadminCatProp.setProperty "SelectionLanguage", "XPath"
set selectedadminCatPropnode = docadminCatProp.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCatPropnodes=docadminCatProp.documentElement.selectNodes("/languages/language")
function getadminCatPropLngStr(instring)
	temp = selectedadminCatPropnode.selectSingleNode(instring).text
	getadminCatPropLngStr = temp
end function
%>

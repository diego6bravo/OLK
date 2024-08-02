<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminDefObjs.xml"
set docadminDefObjs = server.CreateObject("MSXML2.DOMDocument")
docadminDefObjs.async = False
DocadminDefObjs.Load(server.MapPath(xmlfilename)) 
docadminDefObjs.setProperty "SelectionLanguage", "XPath"
set selectedadminDefObjsnode = docadminDefObjs.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminDefObjsnodes=docadminDefObjs.documentElement.selectNodes("/languages/language")
function getadminDefObjsLngStr(instring)
	temp = selectedadminDefObjsnode.selectSingleNode(instring).text
	getadminDefObjsLngStr = temp
end function
%>

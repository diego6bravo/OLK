<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminMyData.xml"
set docadminMyData = server.CreateObject("MSXML2.DOMDocument")
docadminMyData.async = False
DocadminMyData.Load(server.MapPath(xmlfilename)) 
docadminMyData.setProperty "SelectionLanguage", "XPath"
set selectedadminMyDatanode = docadminMyData.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminMyDatanodes=docadminMyData.documentElement.selectNodes("/languages/language")
function getadminMyDataLngStr(instring)
	temp = selectedadminMyDatanode.selectSingleNode(instring).text
	getadminMyDataLngStr = temp
end function
%>

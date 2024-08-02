<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSecIndex.xml"
set docadminSecIndex = server.CreateObject("MSXML2.DOMDocument")
docadminSecIndex.async = False
DocadminSecIndex.Load(server.MapPath(xmlfilename)) 
docadminSecIndex.setProperty "SelectionLanguage", "XPath"
set selectedadminSecIndexnode = docadminSecIndex.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSecIndexnodes=docadminSecIndex.documentElement.selectNodes("/languages/language")
function getadminSecIndexLngStr(instring)
	temp = selectedadminSecIndexnode.selectSingleNode(instring).text
	getadminSecIndexLngStr = temp
end function
%>

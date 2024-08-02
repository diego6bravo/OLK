<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "addItem.xml"
set docaddItem = server.CreateObject("MSXML2.DOMDocument")
docaddItem.async = False
DocaddItem.Load(server.MapPath(xmlfilename)) 
docaddItem.setProperty "SelectionLanguage", "XPath"
set selectedaddItemnode = docaddItem.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaddItemnodes=docaddItem.documentElement.selectNodes("/languages/language")
function getaddItemLngStr(instring)
	temp = selectedaddItemnode.selectSingleNode(instring).text
	getaddItemLngStr = temp
end function
%>

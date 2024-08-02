<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "addCardAddData.xml"
set docaddCardAddData = server.CreateObject("MSXML2.DOMDocument")
docaddCardAddData.async = False
DocaddCardAddData.Load(server.MapPath(xmlfilename)) 
docaddCardAddData.setProperty "SelectionLanguage", "XPath"
set selectedaddCardAddDatanode = docaddCardAddData.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaddCardAddDatanodes=docaddCardAddData.documentElement.selectNodes("/languages/language")
function getaddCardAddDataLngStr(instring)
	temp = selectedaddCardAddDatanode.selectSingleNode(instring).text
	getaddCardAddDataLngStr = temp
end function
%>

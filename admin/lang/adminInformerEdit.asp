<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminInformerEdit.xml"
set docadminInformerEdit = server.CreateObject("MSXML2.DOMDocument")
docadminInformerEdit.async = False
DocadminInformerEdit.Load(server.MapPath(xmlfilename)) 
docadminInformerEdit.setProperty "SelectionLanguage", "XPath"
set selectedadminInformerEditnode = docadminInformerEdit.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminInformerEditnodes=docadminInformerEdit.documentElement.selectNodes("/languages/language")
function getadminInformerEditLngStr(instring)
	temp = selectedadminInformerEditnode.selectSingleNode(instring).text
	getadminInformerEditLngStr = temp
end function
%>

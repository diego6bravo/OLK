<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "activationConfirm.xml"
set docactivationConfirm = server.CreateObject("MSXML2.DOMDocument")
docactivationConfirm.async = False
DocactivationConfirm.Load(server.MapPath(xmlfilename)) 
docactivationConfirm.setProperty "SelectionLanguage", "XPath"
set selectedactivationConfirmnode = docactivationConfirm.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedactivationConfirmnodes=docactivationConfirm.documentElement.selectNodes("/languages/language")
function getactivationConfirmLngStr(instring)
	temp = selectedactivationConfirmnode.selectSingleNode(instring).text
	getactivationConfirmLngStr = temp
end function
%>

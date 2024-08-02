<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCart.xml"
set docadminCart = server.CreateObject("MSXML2.DOMDocument")
docadminCart.async = False
DocadminCart.Load(server.MapPath(xmlfilename)) 
docadminCart.setProperty "SelectionLanguage", "XPath"
set selectedadminCartnode = docadminCart.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCartnodes=docadminCart.documentElement.selectNodes("/languages/language")
function getadminCartLngStr(instring)
	temp = selectedadminCartnode.selectSingleNode(instring).text
	getadminCartLngStr = temp
end function
%>

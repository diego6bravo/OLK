<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminLogos.xml"
set docadminLogos = server.CreateObject("MSXML2.DOMDocument")
docadminLogos.async = False
DocadminLogos.Load(server.MapPath(xmlfilename)) 
docadminLogos.setProperty "SelectionLanguage", "XPath"
set selectedadminLogosnode = docadminLogos.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminLogosnodes=docadminLogos.documentElement.selectNodes("/languages/language")
function getadminLogosLngStr(instring)
	temp = selectedadminLogosnode.selectSingleNode(instring).text
	getadminLogosLngStr = temp
end function
%>

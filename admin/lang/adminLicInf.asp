<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminLicInf.xml"
set docadminLicInf = server.CreateObject("MSXML2.DOMDocument")
docadminLicInf.async = False
DocadminLicInf.Load(server.MapPath(xmlfilename)) 
docadminLicInf.setProperty "SelectionLanguage", "XPath"
set selectedadminLicInfnode = docadminLicInf.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminLicInfnodes=docadminLicInf.documentElement.selectNodes("/languages/language")
function getadminLicInfLngStr(instring)
	temp = selectedadminLicInfnode.selectSingleNode(instring).text
	getadminLicInfLngStr = temp
end function
%>

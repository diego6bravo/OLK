<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminLanguages.xml"
set docadminLanguages = server.CreateObject("MSXML2.DOMDocument")
docadminLanguages.async = False
DocadminLanguages.Load(server.MapPath(xmlfilename)) 
docadminLanguages.setProperty "SelectionLanguage", "XPath"
set selectedadminLanguagesnode = docadminLanguages.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminLanguagesnodes=docadminLanguages.documentElement.selectNodes("/languages/language")
function getadminLanguagesLngStr(instring)
	temp = selectedadminLanguagesnode.selectSingleNode(instring).text
	getadminLanguagesLngStr = temp
end function
%>

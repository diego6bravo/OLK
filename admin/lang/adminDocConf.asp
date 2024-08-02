<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminDocConf.xml"
set docadminDocConf = server.CreateObject("MSXML2.DOMDocument")
docadminDocConf.async = False
DocadminDocConf.Load(server.MapPath(xmlfilename)) 
docadminDocConf.setProperty "SelectionLanguage", "XPath"
set selectedadminDocConfnode = docadminDocConf.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminDocConfnodes=docadminDocConf.documentElement.selectNodes("/languages/language")
function getadminDocConfLngStr(instring)
	temp = selectedadminDocConfnode.selectSingleNode(instring).text
	getadminDocConfLngStr = temp
end function
%>

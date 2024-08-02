<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCUFD.xml"
set docadminCUFD = server.CreateObject("MSXML2.DOMDocument")
docadminCUFD.async = False
DocadminCUFD.Load(server.MapPath(xmlfilename)) 
docadminCUFD.setProperty "SelectionLanguage", "XPath"
set selectedadminCUFDnode = docadminCUFD.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCUFDnodes=docadminCUFD.documentElement.selectNodes("/languages/language")
function getadminCUFDLngStr(instring)
	temp = selectedadminCUFDnode.selectSingleNode(instring).text
	getadminCUFDLngStr = temp
end function
%>

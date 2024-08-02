<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCatOpt.xml"
set docadminCatOpt = server.CreateObject("MSXML2.DOMDocument")
docadminCatOpt.async = False
DocadminCatOpt.Load(server.MapPath(xmlfilename)) 
docadminCatOpt.setProperty "SelectionLanguage", "XPath"
set selectedadminCatOptnode = docadminCatOpt.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCatOptnodes=docadminCatOpt.documentElement.selectNodes("/languages/language")
function getadminCatOptLngStr(instring)
	temp = selectedadminCatOptnode.selectSingleNode(instring).text
	getadminCatOptLngStr = temp
end function
%>

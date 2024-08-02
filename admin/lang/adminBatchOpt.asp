<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminBatchOpt.xml"
set docadminBatchOpt = server.CreateObject("MSXML2.DOMDocument")
docadminBatchOpt.async = False
DocadminBatchOpt.Load(server.MapPath(xmlfilename)) 
docadminBatchOpt.setProperty "SelectionLanguage", "XPath"
set selectedadminBatchOptnode = docadminBatchOpt.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminBatchOptnodes=docadminBatchOpt.documentElement.selectNodes("/languages/language")
function getadminBatchOptLngStr(instring)
	temp = selectedadminBatchOptnode.selectSingleNode(instring).text
	getadminBatchOptLngStr = temp
end function
%>

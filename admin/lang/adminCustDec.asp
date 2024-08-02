<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminCustDec.xml"
set docadminCustDec = server.CreateObject("MSXML2.DOMDocument")
docadminCustDec.async = False
DocadminCustDec.Load(server.MapPath(xmlfilename)) 
docadminCustDec.setProperty "SelectionLanguage", "XPath"
set selectedadminCustDecnode = docadminCustDec.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminCustDecnodes=docadminCustDec.documentElement.selectNodes("/languages/language")
function getadminCustDecLngStr(instring)
	temp = selectedadminCustDecnode.selectSingleNode(instring).text
	getadminCustDecLngStr = temp
end function
%>

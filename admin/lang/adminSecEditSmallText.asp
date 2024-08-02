<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSecEditSmallText.xml"
set docadminSecEditSmallText = server.CreateObject("MSXML2.DOMDocument")
docadminSecEditSmallText.async = False
DocadminSecEditSmallText.Load(server.MapPath(xmlfilename)) 
docadminSecEditSmallText.setProperty "SelectionLanguage", "XPath"
set selectedadminSecEditSmallTextnode = docadminSecEditSmallText.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSecEditSmallTextnodes=docadminSecEditSmallText.documentElement.selectNodes("/languages/language")
function getadminSecEditSmallTextLngStr(instring)
	temp = selectedadminSecEditSmallTextnode.selectSingleNode(instring).text
	getadminSecEditSmallTextLngStr = temp
end function
%>

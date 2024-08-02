<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminPolls.xml"
set docadminPolls = server.CreateObject("MSXML2.DOMDocument")
docadminPolls.async = False
DocadminPolls.Load(server.MapPath(xmlfilename)) 
docadminPolls.setProperty "SelectionLanguage", "XPath"
set selectedadminPollsnode = docadminPolls.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminPollsnodes=docadminPolls.documentElement.selectNodes("/languages/language")
function getadminPollsLngStr(instring)
	temp = selectedadminPollsnode.selectSingleNode(instring).text
	getadminPollsLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "transComplete.xml"
set doctransComplete = server.CreateObject("MSXML2.DOMDocument")
doctransComplete.async = False
DoctransComplete.Load(server.MapPath(xmlfilename)) 
doctransComplete.setProperty "SelectionLanguage", "XPath"
set selectedtransCompletenode = doctransComplete.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedtransCompletenodes=doctransComplete.documentElement.selectNodes("/languages/language")
function gettransCompleteLngStr(instring)
	temp = selectedtransCompletenode.selectSingleNode(instring).text
	gettransCompleteLngStr = temp
end function
%>

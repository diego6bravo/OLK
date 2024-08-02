<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "transError.xml"
set doctransError = server.CreateObject("MSXML2.DOMDocument")
doctransError.async = False
DoctransError.Load(server.MapPath(xmlfilename)) 
doctransError.setProperty "SelectionLanguage", "XPath"
set selectedtransErrornode = doctransError.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedtransErrornodes=doctransError.documentElement.selectNodes("/languages/language")
function gettransErrorLngStr(instring)
	temp = selectedtransErrornode.selectSingleNode(instring).text
	gettransErrorLngStr = temp
end function
%>

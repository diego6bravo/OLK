<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "carddetails.xml"
set doccarddetails = server.CreateObject("MSXML2.DOMDocument")
doccarddetails.async = False
Doccarddetails.Load(server.MapPath(xmlfilename)) 
doccarddetails.setProperty "SelectionLanguage", "XPath"
set selectedcarddetailsnode = doccarddetails.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcarddetailsnodes=doccarddetails.documentElement.selectNodes("/languages/language")
function getcarddetailsLngStr(instring)
	temp = selectedcarddetailsnode.selectSingleNode(instring).text
	getcarddetailsLngStr = temp
end function
%>

<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "delOrderCheckSerial.xml"
set docdelOrderCheckSerial = server.CreateObject("MSXML2.DOMDocument")
docdelOrderCheckSerial.async = False
DocdelOrderCheckSerial.Load(server.MapPath(xmlfilename)) 
docdelOrderCheckSerial.setProperty "SelectionLanguage", "XPath"
set selecteddelOrderCheckSerialnode = docdelOrderCheckSerial.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selecteddelOrderCheckSerialnodes=docdelOrderCheckSerial.documentElement.selectNodes("/languages/language")
function getdelOrderCheckSerialLngStr(instring)
	temp = selecteddelOrderCheckSerialnode.selectSingleNode(instring).text
	getdelOrderCheckSerialLngStr = temp
end function
%>

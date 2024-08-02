<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "updateDb.xml"
set docupdateDb = server.CreateObject("MSXML2.DOMDocument")
docupdateDb.async = False
DocupdateDb.Load(server.MapPath(xmlfilename)) 
docupdateDb.setProperty "SelectionLanguage", "XPath"
set selectedupdateDbnode = docupdateDb.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedupdateDbnodes=docupdateDb.documentElement.selectNodes("/languages/language")
function getupdateDbLngStr(instring)
	temp = selectedupdateDbnode.selectSingleNode(instring).text
	getupdateDbLngStr = temp
end function
%>

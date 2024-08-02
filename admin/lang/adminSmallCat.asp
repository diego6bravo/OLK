<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminSmallCat.xml"
set docadminSmallCat = server.CreateObject("MSXML2.DOMDocument")
docadminSmallCat.async = False
DocadminSmallCat.Load(server.MapPath(xmlfilename)) 
docadminSmallCat.setProperty "SelectionLanguage", "XPath"
set selectedadminSmallCatnode = docadminSmallCat.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminSmallCatnodes=docadminSmallCat.documentElement.selectNodes("/languages/language")
function getadminSmallCatLngStr(instring)
	temp = selectedadminSmallCatnode.selectSingleNode(instring).text
	getadminSmallCatLngStr = temp
end function
%>

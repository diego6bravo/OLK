<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "adminClientsPriceList.xml"
set docadminClientsPriceList = server.CreateObject("MSXML2.DOMDocument")
docadminClientsPriceList.async = False
DocadminClientsPriceList.Load(server.MapPath(xmlfilename)) 
docadminClientsPriceList.setProperty "SelectionLanguage", "XPath"
set selectedadminClientsPriceListnode = docadminClientsPriceList.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedadminClientsPriceListnodes=docadminClientsPriceList.documentElement.selectNodes("/languages/language")
function getadminClientsPriceListLngStr(instring)
	temp = selectedadminClientsPriceListnode.selectSingleNode(instring).text
	getadminClientsPriceListLngStr = temp
end function
%>

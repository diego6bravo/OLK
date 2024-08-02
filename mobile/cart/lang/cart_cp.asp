<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cart_cp.xml"
set doccart_cp = server.CreateObject("MSXML2.DOMDocument")
doccart_cp.async = False
Doccart_cp.Load(server.MapPath(xmlfilename)) 
doccart_cp.setProperty "SelectionLanguage", "XPath"
set selectedcart_cpnode = doccart_cp.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcart_cpnodes=doccart_cp.documentElement.selectNodes("/languages/language")
function getcart_cpLngStr(instring)
	temp = selectedcart_cpnode.selectSingleNode(instring).text
	getcart_cpLngStr = temp
end function
%>

<script language="VB" runat="server"> 
Dim myNode As XmlNode 
Private Sub loadLanguage() 
	Dim myLng As String = "en" 
	If (not Request.Cookies("myLng") is Nothing) Then myLng = Request.Cookies("myLng").Value 
    Dim xmlFileName As String = Server.MapPath(Request.ServerVariables("SCRIPT_NAME").Replace("update.aspx", "lang/update.xml")) 
    Dim oDoc As XmlDocument = new XmlDocument() 
    oDoc.Load(xmlFileName) 
    Dim oManager As XmlNamespaceManager = new XmlNamespaceManager(oDoc.NameTable) 
    oManager.AddNamespace("SelectionLanguage", "XPath") 
    myNode = oDoc.SelectSingleNode("/languages/language[@xml:lang='" + getLng() + "']", oManager) 
End Sub 
Private Function getLng() As String 
  Dim myLng As String = "en" 
  If (not Request.Cookies("myLng") is Nothing) Then myLng = Request.Cookies("myLng").Value 
  return myLng 
End Function 
Private Function getLangVal(ByVal key As String) As String 
	return myNode.SelectSingleNode(key).InnerText 
End Function 
</script> 

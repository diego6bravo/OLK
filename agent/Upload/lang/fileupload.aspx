<script language="C#" runat="server"> 
XmlNode myNode; 
private void loadLanguage() 
{ 
	string myLng = "en"; 
	if (Request.Cookies["myLng"] != null) myLng = Request.Cookies["myLng"].Value; 
    string xmlFileName = Server.MapPath(Request.ServerVariables["SCRIPT_NAME"].Replace("fileupload.aspx", "lang/fileupload.xml")); 
    XmlDocument oDoc = new XmlDocument(); 
    oDoc.Load(xmlFileName); 
    XmlNamespaceManager oManager = new XmlNamespaceManager(oDoc.NameTable); 
    oManager.AddNamespace("SelectionLanguage", "XPath"); 
    myNode = oDoc.SelectSingleNode("/languages/language[@xml:lang='" + getLng() + "']", oManager); 
} 
private string getLng() { 
string myLng ="en"; 
if (Request.Cookies["myLng"] != null) myLng = Request.Cookies["myLng"].Value; return myLng; 
} 
private string getLangVal(string key) 
{ 
	return myNode.SelectSingleNode(key).InnerText; 
} 
</script> 

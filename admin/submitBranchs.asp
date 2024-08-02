<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="myHTMLEncode.asp"-->
<!--#include file="adminTradSave.asp"-->
<%
Redir = request("redir")
           	set rs = server.createobject("ADODB.RecordSet")
           	Select Case Request("cmd") 
           		Case "addBranch"
           			If Request("Active") = "Y" Then ActiveBranch = "Y" Else ActiveBranch = "N"
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = adCmdStoredProc
					cmd.CommandText = "DBOLKAdminBranchs" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@BranchName") = saveHTMLDecode(Request("BranchName"), True)
					cmd("@WhsCode") = saveHTMLDecode(Request("WhsCode"), True)
					cmd("@Active") = ActiveBranch
					cmd.execute()
					
					branchIndex = cmd("@BranchIndex")
					If Request("BranchNameTrad") <> "" Then
						SaveNewTrad Request("BranchNameTrad"), "Branchs", "branchIndex", "alterBranchName", branchIndex
					End If
	   			Case "delBranch"
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = adCmdStoredProc
					cmd.CommandText = "DBOLKAdminBranchs" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@BranchIndex") = Request("branchIndex")
					cmd("@Action") 		= "D"
					cmd.execute()
	   			Case "activeBranch"
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = adCmdStoredProc
					cmd.CommandText = "DBOLKAdminBranchs" & Session("ID")
					cmd.Parameters.Refresh()
					
	   				GetQuery rs, 5, null, null
	   				do while not rs.eof
	   					If Request("Active" & rs("branchIndex")) = "Y" Then ActiveBranch = "Y" Else ActiveBranch = "N"
	   					cmd("@BranchIndex") = rs("BranchIndex")
	   					cmd("@BranchName") 	= saveHTMLDecode(Request("BranchName" & rs("BranchIndex")), True)
	   					cmd("@Active") 		= ActiveBranch
	   					cmd.execute()
	   				rs.movenext
	   				loop
	   			Case "updateBranch"
	   				GetQuery rs, 7, null, null
	   				
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = adCmdStoredProc
					cmd.CommandText = "DBOLKAdminBranchsCards" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@BranchIndex") = Request("branchIndex")
			   		do while not rs.eof 
			   			cmd("@CreditCard") 		= rs("CreditCard")
			   			cmd("@AcctCode") 		= Request("CreditAcctCode" & rs("CreditCard"))	
			   			cmd("@OIRAcctCode") 	= Request("OIRCreditAcctCode" & rs("CreditCard"))
			   			cmd.execute()
			   		rs.movenext
			   		loop
		            If Request("Active") = "Y" Then ActiveBranch = "Y" Else ActiveBranch = "N"
			   		
					set cmd = Server.CreateObject("ADODB.Command")
					cmd.ActiveConnection = connCommon
					cmd.CommandType = adCmdStoredProc
					cmd.CommandText = "DBOLKAdminBranchs" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@BranchIndex") 	= Request("branchIndex")
					cmd("@BranchName")		= saveHTMLDecode(Request("branchName"), True)
					cmd("@WhsCode")			= saveHTMLDecode(Request("WhsCode"), True)
					cmd("@Active") 			= ActiveBranch
					If Request("OQUTSeries") <> "" Then 	cmd("@OQUTSeries")		= Request("OQUTSeries")
					If Request("OINVSeries") <> "" Then 	cmd("@OINVSeries")		= Request("OINVSeries")
					If Request("OINVResSeries") <> "" Then 	cmd("@OINVResSeries")	= Request("OINVResSeries")
					If Request("ORDRSeries") <> "" Then 	cmd("@ORDRSeries")		= Request("ORDRSeries")
					If Request("ORCTSeries") <> "" Then 	cmd("@ORCTSeries")		= Request("ORCTSeries")
					If Request("ODLNSeries") <> "" Then 	cmd("@ODLNSeries")		= Request("ODLNSeries")
					If Request("OIRISeries") <> "" Then 	cmd("@OIRISeries")		= Request("OIRISeries")
					If Request("OIRRSeries") <> "" Then 	cmd("@OIRRSeries")		= Request("OIRRSeries")
					cmd("@CashAcct")		= Request("CashAcct")
					cmd("@CheckAcct")		= Request("CheckAcct")
					cmd("@OIRCashAcct")		= Request("OIRCashAcct")
					cmd("@OIRCheckAcct")	= Request("OIRCheckAcct")
					cmd("@Action") = "U"
		            
					cmd.execute()
					
		            If Request("btnApply") <> "" Then
						branchIndex = "&branchIndex=" & Request("branchIndex")
		            	Redir = "adminBranchsEdit.asp"
		            Else
		            	Redir = "adminBranchs.asp"
		            End If
            	Case "addPos"
		            sql = "declare @posIndex int set @posIndex = " & _
		                  "ISNULL((select max(posIndex)+1 from olkbranchspos where branchIndex = " & Request("branchIndex") & "),0) " & _
		                  "insert olkbranchspos values(" & Request("branchIndex") & ", @posIndex, '" & Request("posName") & "')"
		                        	  branchIndex = "&branchIndex=" & Request("branchIndex") & "&place=p#posName"
            	Case "delPOS"
		            sql = "update olkBranchsPOS set Active = 'D' where posIndex = " & Request("posIndex") & " and branchIndex = " & Request("branchIndex")
		                        	  branchIndex = "&branchIndex=" & Request("branchIndex") & "&place=a#posName"
	  			End Select
	  'conn.execute(sql)
	  set rs = nothing
	  conn.close
	  if Redir <> "" then response.redirect Redir & "?1=1" & branchIndex
	          %>

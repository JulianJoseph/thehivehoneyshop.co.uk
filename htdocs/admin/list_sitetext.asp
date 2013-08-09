<%@ Language=VBScript %>
<!--#include file="adminutils.asp"-->
<%
	If Session("LoggedIn") <> Session("PassWord") Then
		Response.Redirect("default.asp")
	End If
		'On Error Resume Next
		response.Expires=-1000
		Dim oRS
		Dim sSQL
		Dim fo
		Dim sConnStr
		Dim oFile
		Dim sPath
		Dim Q
		Dim fld
		Dim i
		Dim s
		Dim SortBy

		
		Q = Chr(34)
	

		sSQL = "SELECT * FROM SITETEXT ORDER BY ID"
		Set oRS=Server.CreateObject("ADODB.Recordset")		
		oRS.Open sSQL, gsConnStr, adOpenDynamic, adLockOptimistic
		'======================================================================			
		%>
		<html>
		<head>
			<TITLE>The Hive Honey Shop: Web Site Administration: SiteText Administration</TITLE>
			<LINK rel="stylesheet" href="hive.css" type="text/css">
		</head>
		<body>
		<!--#include file="header.asp"-->
		<h3>Website Text Administration</h3>
		<FORM action="mainpage.asp" method="post" ID="Form1">
	
			<table border="1" cellspacing="0" cellpadding="2"><thead><tr>
						<td width="180px"><b>Entry</b></td>
				<td width="440px" ><b>Current Live Text</b></td>
				</tr></thead>
				<tbody>
		<% If Not oRS.BOF And Not oRS.EOF Then
			oRS.MoveFirst 
			While Not oRS.EOF
				%>
				<tr>
				<td><%=oRS.Fields("Description")%><br /> <a href="update_sitetext.asp?id=<%=oRS.Fields("ID")%>">Edit</a></td>
				
				<td>
				<%=oRS.Fields("SiteText")%>
				</td>
				</tr>
				<%
				oRS.MoveNext 
			Wend
		End If 
			%>
			<tr><td colspan=4>
				<input type=button value="Main Menu" onclick="window.location.href='mainpage.asp';">
				<input type="button" value="Log Out" ID=btnLogOut NAME=btnLogOut onclick="window.location.href='logout.asp';">
			</td></tr>
			</FORM>
			<tbody></table>
			<%
		'======================================================================	
		oRS.Close
		Set oRS=Nothing

%>
	</form>
	</body>
</html>

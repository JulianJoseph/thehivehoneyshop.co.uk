<%@ Language=VBScript %>
<!--#include file="adminutils.asp"-->
<%
	Response.Expires = -1000 'Makes the browser not cache this page
	Response.Buffer = True 'Buffers the content so our Response.Redirect will work
	'check permissions
	If (Session("LoggedIn") <> Session("PassWord")) Or (Session("LoggedIn") = "") Then
		Response.Redirect("default.asp")
	End If

	Dim sSQL
	Dim oRS
	Dim Q
	Dim s
	
	
	sSQL = "SELECT * FROM Category WHERE CategoryID = " & request.QueryString("CategoryID") & " ORDER BY CategoryID"

%>
<html>
<head>
	<title>The Hive Honey Shop: Web Site Administration</title>
	<LINK rel="stylesheet" href="hive.css" type="text/css">
</head>
<body>
<!--#include file="header.asp"-->

<%

	Set oRS=Server.CreateObject("ADODB.Recordset")
	oRS.Open sSQL, gsConnStr
	If Not oRS.BOF And Not oRS.EOF Then %>
		<form  method="post" action="updatecategory.asp" id="form1">
		<table border="1" cellpadding="1" cellspacing="1">
			<tr>
				<td>Category Description:&nbsp;</td>
				<td><input type="text" id=txtDescription name=txtDescription value="<%=oRS("Description")%>" size="<%=oRS("Description").DefinedSize%>"></td>
				
			</tr>
			<tr>
				<td>Header Text:&nbsp;</td>
				<td><TEXTAREA id=txtHeaderText name=txtHeaderText cols="75" rows="15"><%=oRS("HeaderText")%></textarea></td>
			</tr>
			<tr>
				<td>Footer Text:&nbsp;</td>
				<td><TEXTAREA name=txtFooterText id=txtFooterText cols="75" rows="15"><%=oRS("FooterText")%></textarea></td>
			</tr>
			<tr>
			<td>
				<input type="submit" value="Save" ID=btnSave NAME=btnSave>&nbsp;
				<input type="reset" value="Cancel" ID=btnCancel NAME=btnCancel>&nbsp;
				<input type="button" value="Close" ID=btnClose NAME=btnClose onclick="window.location.href='mainpage.asp';">
			</td>
			</tr>
		</table>
		<input type="hidden" value="editcategory" name=editcat id=editcat>
		<input type="hidden" name=CategoryID id=CategoryID value="<%=oRS("CategoryID")%>">
		</form>
	<%End If
	oRS.Close
	Set oRS = Nothing

%>
<p><a href="logout.asp">Log Out</a></p>
</body>
</html>
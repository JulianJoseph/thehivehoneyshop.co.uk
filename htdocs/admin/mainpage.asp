<%@ Language=VBScript %>
<!--#include file="adminutils.asp"-->
<%
	Response.Expires = -1000 'Makes the browser not cache this page
	Response.Buffer = True 'Buffers the content so our Response.Redirect will work
	If (Session("LoggedIn") <> Session("PassWord")) Or (Session("LoggedIn") = "") Then
		Response.Redirect("default.asp")
	End If


	Dim sSQL
	Dim oRS
	'Dim sConnStr
	Dim Q
	Dim s
	
	Q = Chr(34)

	sSQL = "SELECT CategoryID, Description FROM Category ORDER BY Description"

%>
<html>
<head>
	<title>The Hive Honey Shop: Web Site Administration</title>
	<LINK rel="stylesheet" href="hive.css" type="text/css">
</head>
<body>
	<!--#include file="header.asp"-->
	<h3>Product Categories</h3>
	<p>Click on a category title to see a list of products or click <i>Edit</i> to edit Category details</p>

<%

	Set oRS=Server.CreateObject("ADODB.Recordset")
	oRS.Open sSQL, gsConnStr
	If Not oRS.BOF And Not oRS.EOF Then %>
		<form id=form1 method="post" action="updatecategory.asp">
		<TABLE BORDER="1" CELLSPACING="1" CELLPADDING="1">
			<tr>
				<td><i>&nbsp;</i></td>
				<td><i>Description</i></td>
				<td><i>Delete?</i></td>
			</tr>

		<%While Not oRS.EOF
			%>
			<tr>
				<td><a href="EditCategory.asp?CategoryID=<%=oRS("CategoryID")%>">Edit</a></td>
				<td><a href="list_products.asp?categoryid=<%=oRS.Fields("CategoryID")%>"><%=oRS.Fields("Description")%></a></td>
				<td><input type="checkbox" name=delCategory_<%=oRS("CategoryID")%>></td>
			</tr>
			<%
			oRS.MoveNext
		Wend
		%>
			<tr>
				<td>New</td>
				<td><input type="text" Name=NewCategory ID=NewCategory size="<%=oRS("Description").DefinedSize%>"></td>
				<td>&nbsp;</td>
			</tr>
		
		</TABLE>
		<input type="submit" value="Save Product Category Changes">
<%		
' Check for online site order status
	Dim sSiteOnlineEnabled, sSiteOnlineBtnText, sSiteOnlineStatus
	sSiteOnlineEnabled = 0
	sSQL = "SELECT OnlineOrderEnabled FROM admin where id=1"
	Set oRS=Server.CreateObject("ADODB.Recordset")
	oRS.Open sSQL, gsConnStr
	If Not oRS.BOF And Not oRS.EOF Then	sSiteOnlineEnabled = oRS("OnlineOrderEnabled")
sSiteOnlineStatus = "[ Online Site Orders "
	If sSiteOnlineEnabled = True Then 
		sSiteOnlineStatus= sSiteOnlineStatus & "Enabled ]"
		sSiteOnlineBtnText = "Disable Online Orders"
		sSiteOnlineStatusColour = "green"
		Else
		sSiteOnlineStatus = sSiteOnlineStatus & "Disabled ]"
		sSiteOnlineBtnText = "Enable Online Orders"
		sSiteOnlineStatusColour="red"
	End If
%>
	<br><br>
	<p>
<table  BORDER="1"><tr><td>	<a href="logout.asp">Log Out</a>&nbsp; </td>
	<td><a href="list_sitetext.asp">Edit website text content</a> &nbsp; </td>
	<td><a href="reportadminaudit.asp">Database Product Audit Report</a> &nbsp; </td>

	<td style="background-color:<%=sSiteOnlineStatusColour %>;"><%=sSiteOnlineStatus %><br>
		<input type="submit" value="<%=sSiteOnlineBtnText %>" ID=btnOnlineOrders NAME=btnOnlineOrders ></td></tr></table>
	</p>
</form>
		<%
	End If
	oRS.Close
	Set oRS = Nothing
	
	If Request.QueryString("Action") = "delete_forbidden" Then
		%>
		<script>
			alert("You may not delete this category as it still contains products.\n You must delete or move all its products before you delete it.");
			document.location.href = "mainpage.asp";
		</script>
		<%
	End If	

%>
</body>
</html>

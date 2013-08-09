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
		Dim sDescription
		
		Q = Chr(34)
	
		sSQL = "SELECT Description FROM Category WHERE CategoryID = " & Request.QueryString("categoryid")
		Set oRS=Server.CreateObject("ADODB.Recordset")		
		oRS.Open sSQL, gsConnStr
		If Not oRS.BOF and Not oRS.EOF Then
			sDescription = oRS("Description")
		Else
			sDescription = "[Category Description Missing]"
		End If
		oRS.Close
		Set oRS = Nothing
		
		If request.QueryString("SortBy") <> "" Then
			SortBy = request.QueryString("SortBy")
		Else
			SortBy = "Title"
		End If
		sSQL = "SELECT * FROM qAdminProducts WHERE CategoryID = " & Request.QueryString("categoryid") & " ORDER BY " & SortBy 
		Set oRS=Server.CreateObject("ADODB.Recordset")		
		oRS.Open sSQL, gsConnStr, adOpenDynamic, adLockOptimistic
		'======================================================================			
		%>
		<html>
		<head>
			<TITLE>The Hive Honey Shop: Web Site Administration</TITLE>
			<LINK rel="stylesheet" href="hive.css" type="text/css">
		</head>
		<body>
		<!--#include file="header.asp"-->
		<h2><%=sDescription%></h2>
		<p><a href="list_products.asp?SortBy=Title&CategoryID=<%=Request.QueryString("categoryid")%>">[Sort by Description]</a><a href="list_products.asp?SortBy=ProductCode&CategoryID=<%=Request.QueryString("categoryid")%>">[Sort by Product Code]</a></p>
		
		<FORM action="updateproduct.asp?categoryid=<%=Request.QueryString("categoryid")%>" method="post" ID="Form1">
	
			<table border="1" cellspacing="0" cellpadding="2"><thead><tr>
			<td>&nbsp;</td>
				<td><i>Product Description</i></td>
				<td><i>Product Code</i></td>
				<td><i>Discontinued?</i></td>
				<td><i>Publish?</i></td>
				<td><i>Delete?</i></td>
				</tr></thead>
				<tbody>
		<% If Not oRS.BOF And Not oRS.EOF Then
			oRS.MoveFirst 
			While Not oRS.EOF
				%>
				<tr>
				<td><a href="update_catalogue.asp?productid=<%=oRS.Fields("ProductID")%>&categoryid=<%=Request.QueryString("categoryid")%>">Edit</a>
				<td><%=oRS("Title")%>&nbsp;</td>
				<td><%=oRS("ProductCode")%>&nbsp;</td>
				<td><%If (oRS("Discontinued")) Then Response.Write("Y") Else Response.Write("N") End If%>&nbsp;</td>
				<td><%If (oRS("Publish")) Then Response.Write("Y") Else Response.Write("N") End If%>&nbsp;</td>
				<td><input type=checkbox id=chkDelete<%=oRS.Fields("ProductID")%> name=chkDelete<%=oRS.Fields("ProductID")%>></td>
				</tr>
				<%
				oRS.MoveNext 
			Wend
		End If 
			%>
			<tr><td>New</td><td><input type=text name=producttitle id=producttitle size=50></td><td>&nbsp;</td></tr>
			<tr><td colspan=4>
				<input type=submit value=Save>
				<input type=reset value=Cancel>
				<input type=button value=Close onclick="window.location.href='mainpage.asp';">
			</td></tr>
			</FORM>
			<tbody></table>
			<%
		'======================================================================	
		oRS.Close
		Set oRS=Nothing

%>
	</form>
	<p><a href="logout.asp">Log Out</a></p>
	</body>
</html>

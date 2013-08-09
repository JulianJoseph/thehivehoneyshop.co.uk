<%@ Language=VBScript %>
<!--#include file="adminutils.asp"-->
<!--#include file="FCKeditor/fckeditor.asp" -->
<html>
<head>
	<title>The Hive Honey Shop: Web Site Administration</title>
	<LINK rel="stylesheet" href="hive.css" type="text/css">
	<script type="text/javascript" src="FCKeditor/fckeditor.js"></script>
	
	<script type="text/javascript">
	window.onload = function()
	{
		var sBasePath = "FCKeditor/"; 
		var oFCKeditor = new FCKeditor( 'txtSiteText' ) ;
		oFCKeditor.BasePath	= sBasePath ;
		oFCKeditor.ToolbarSet = "Hive";
		oFCKeditor.ReplaceTextarea() ;
	}
	</script></head>
<body>

<%
	Response.Expires = -1000 'Makes the browser not cache this page
	Response.Buffer = True 'Buffers the content so our Response.Redirect will work
	Dim sSQL
	Dim oRS
	Dim Q
	Dim s
	Dim sConnStr
	Dim sUpdatedSiteText
	Q = Chr(39)

	'check permissions
	If (Session("LoggedIn") <> Session("PassWord")) Or (Session("LoggedIn") = "") Then
		Response.Redirect("default.asp")
	End If

	' Check if a postback, if so save and then go back to list
	If request.Form("btnPreview")="Preview Unpublished Changes" Or request.Form("btnPublish")="Publish Changes" Then 
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open gsConnStr

		sUpdatedSiteText = Replace(request.Form("txtSiteText") , Chr(39), "''")
		If sUpdatedSiteText="" Then 
			sUpdatedSiteText=" "
		End if

		sSQL = "Update SiteText SET Description = " & Q & request.Form("txtDescription") & Q & " ,SiteTextPreview = " & Q & sUpdatedSiteText & Q & ", LastUpdated=now(), Published='No' WHERE ID = " & request.Form("SiteTextID") 
		oConn.Execute sSQL
'		response.Write "<P>" & sSQL & "</P>"

		oConn.Close 
		'<input type="text" id=txtSiteText name=txtSiteText value="<%=oRS("SiteText")% >"  size="<%=oRS("SiteText").DefinedSize% >">
	Else
		' There are unpublished changes from last time need to be reset to match the current live output
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open gsConnStr
		sSQL = "Update SiteText SET SiteTextPreview = SiteText,Published='Yes' WHERE ID=" & request.QueryString("ID")
		oConn.Execute sSQL
		oConn.Close 
	End If

	If request.Form("btnPublish")="Publish Changes" Then 
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open gsConnStr
		sSQL = "Update SiteText SET SiteText = SiteTextPreview,Published='Yes' WHERE ID=" & request.Form("SiteTextID") 
		oConn.Execute sSQL
		oConn.Close 
		' If published then redirect back to list page
		response.Redirect "list_sitetext.asp" 
	End If

	%>
	<!--#include file="header.asp"-->
	<h3>Website Text Administration > Edit Entry</h3>
	<%	

	' Fetch all values from table for current entry
	sSQL = "SELECT * FROM SiteText WHERE ID = " & request.QueryString("ID")
	Set oRS=Server.CreateObject("ADODB.Recordset")
	oRS.Open sSQL, gsConnStr
	If Not oRS.BOF And Not oRS.EOF Then
	sSiteTextPreview = oRS("SiteTextPreview")
	%>
		<form  method="post" action="update_sitetext.asp?ID=<%=request.QueryString("ID")%>" id="form1">
		<input type="hidden" name=SiteTextID id=SiteTextID value="<%=oRS("ID")%>">
		<p>Website Content Description:&nbsp;<br />
		<input type="text" id=txtDescription name=txtDescription value="<%=oRS("Description")%>" size="<%=oRS("Description").DefinedSize%>"></p>

		<p>Website Content HTML:&nbsp;<br />
		<textarea name="txtSiteText" rows="10" cols="60" style="width: 760px; height: 400px"><%=sSiteTextPreview %></textarea>

		<input type="submit" value="Preview Unpublished Changes" ID=btnPreview NAME=btnPreview>&nbsp;
		<input type="submit" value="Publish Changes" ID=btnPublish NAME=btnPublish>&nbsp;
		<input type="reset" value="Cancel Changes" ID=btnCancel NAME=btnCancel onclick="window.location.href='list_sitetext.asp';">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="Main Menu" ID=btnClose NAME=btnClose onclick="window.location.href='mainpage.asp';">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="Log Out" ID=btnLogOut NAME=btnLogOut onclick="window.location.href='logout.asp';">
		</form>
		
		<HR><%
		' Generate Preview if required
		If oRS("Published")="No" Then  
			response.Write "<b><font color=#ff0000>Unpublished Changes</font></b><HR>" & sSiteTextPreview &"<HR>"
		End if

		response.Write "<b>Current Published Content</b><HR>"
		response.Write oRS("SiteText")
		response.Write "<HR>"
	End If
	oRS.Close
	Set oRS = Nothing
%>
</body>
</html>
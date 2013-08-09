<%@ Language=VBScript %>
<!--#include file="adminutils.asp"-->
<%
Response.Expires = -1000 'Makes the browser not cache this page
Response.Buffer = True 'Buffers the content so our Response.Redirect will work
Session("LoggedIn") = ""
%>
<HTML>
<HEAD>
	<TITLE>The Hive Honey Shop: Web Site Administration</TITLE>
	<LINK rel="stylesheet" href="hive.css" type="text/css">
</HEAD>
<BODY>
<%
If Request.Form("login") = "true" Then
    CheckLogin
Else
    ShowLogin
End If
%>

</BODY>
</HTML>

<%
Sub getPermissions
	Dim rs
	Dim sSQL
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	sSQL = "SELECT UserName, PassWord FROM Security"
	rs.Open sSQL, gsConnStr
	If Not rs.BOF And Not rs.EOF Then
		Session("UserName") = rs("UserName")
		Session("PassWord") = rs("PassWord")
	End If
	rs.Close
	Set rs = Nothing
	
End Sub

Sub CheckLogin
	Call getPermissions()
	If ((Request.Form("passwd")) = Session("PassWord")) And ((Request.Form("username")) = Session("UserName")) Then
		Session("LoggedIn") = Session("PassWord")
		Response.Redirect "mainpage.asp"
	Else
		ShowLogin
		Response.Write("<H2>Login Failed!</H2>")
	End If
End Sub

sub ShowLogin %>
	<P>
	<FORM NAME=frmLogIn ACTION="default.asp" METHOD="POST">
	<H2>The Hive Honey Shop: Web Site Administration</H2>
	<P>
	<TABLE WIDTH=50%>
	  <TR>
		<TD>User Name:
	    </TD>
		<TD><INPUT TYPE=text NAME=username ID=username>
	    </TD>
	  </TR>
	  <TR>
		<TD>Password:
	    </TD>
		<TD><INPUT TYPE=password NAME=passwd ID=passwd>
	    	<INPUT TYPE=hidden NAME=login ID=login VALUE=true>
	    </TD>
	  </TR>
	  <TR>
		<TD>
			<INPUT type="submit" value="Log In">
	    </TD>
	  </TR>
	</TABLE>
	</FORM>
<%end sub%>


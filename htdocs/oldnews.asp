<%@ Language=VBScript%>
<!--#include file="../private/db/connection.asp"-->
<%
	Dim s
	Dim Q
	Dim objXML
	Dim xmlDoc
	Dim oNodeList
	Dim oNode
	Dim i
	Dim sTarget
	
	Q = Chr(34)

	Set objXML = Server.CreateObject("Microsoft.XMLDOM")
	sPath = Server.MapPath("./news/news.xml")
	objXML.async=False
	objXML.load(sPath)	
%>
<html>
<head>
	<title>The Hive Honey Shop : News</title>
	<style>
		A {text-decoration:none}
		A:hover {text-decoration:underline}
	</style>
	<!--#include file="include/javascript.asp"-->
</head>
<!--#include file="include/body.htm"-->
<script language="JavaScript1.2">fwLoadMenus();</script>
<div align="center">

	<!--contact details-->
	<!--#include file="include/contact.htm"-->
	
	<!--masthead-->
	<!--#include file="include/masthead.htm"-->

	<table style="border:solid;border-width:1px" cellpadding="5" cellspacing="0" width="760" bordercolor=#6D1746 bordercolorlight=#6D1746 bordercolordark=#6D1746>
	<tr>
		<td>
					<font face="verdana,arial,sans-serif" color="#6D1746">
					<h3>News</h3>
					</font>
					<font face="verdana,arial,sans-serif" size="2" color="#000000">				
					<p>This page contains links to international news items featuring The Hive Honey Shop, as well as news and updates
						on courses, features and promotions.
					<hr noshade color=#CCCCCC size="1">
						<%
							Dim oPara
							Set oNodeList = objXML.childNodes(1).childNodes
							For Each oNode in oNodeList
								If Left(oNode.childNodes.item(3).text,4) = "http" Then
									sTarget = " target=" & Q & "_blank" & Q
								Else
									sTarget = " "
								End If
								
								s = ""
								s = "<B><A href=" & Q & oNode.childNodes.item(3).text & Q & sTarget & ">" & oNode.childNodes.item(0).text & "&nbsp;</A></B>" 								
								If oNode.childNodes.item(1).text <> "" Then
									s = s & "<I>(" & oNode.childNodes.item(1).text & ")</I>"
								Else
									's = s & "</P>"
								End If
								s = s & "</BR>"
								For Each oPara in oNode.childNodes.item(2).childNodes
									's = s & "<P>" & oPara.text & "</P>"
									s = s & "<BR>" & oPara.text
								Next 
								's = s & "<P><A href=" & Q  & oNode.childNodes.item(3).text & Q & " target=" & Q & "_blank" & Q & ">" & _
								'		oNode.childNodes.item(3).text & "</A></P>"
								s = s & "<HR noshade color=#CCCCCC size=""1"">"
								Response.Write s & vbCrLf
							Next
							Set oPara = Nothing
							Set oNode = Nothing
							Set oNodeList = Nothing
							Set objXML = Nothing
						%>
			</font>
		</td>
	</tr>
	<tr>
		<td colspan="2" align="right"><img src="images/2bees.gif"></td>
	</tr>
	</table>
</div>
</body>
</html>
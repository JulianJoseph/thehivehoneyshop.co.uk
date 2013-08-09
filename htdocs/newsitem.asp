<%@ Language=VBScript%>
<%
	Dim s
	Dim Q
	Dim objXML
	Dim xmlDoc
	Dim oNodeList
	Dim oNode
	Dim i
	
	Q = Chr(34)

	Set objXML = Server.CreateObject("Microsoft.XMLDOM")
	sPath = Server.MapPath("./news/story" & Request.QueryString("newsitem") & ".xml")
	objXML.async=False
	objXML.load(sPath)	
	'Response.Write "<P>XML " & objXML.XML
	'Response.Write "<P>XML " & objXML.childNodes(1).childNodes.length
%>
<html>
<head>
	<title>The Hive Honey Shop : News</title>
	<style>
		A {text-decoration:none}
		A:hover {text-decoration:underline}
	</style>
	<!--#include file="include/javascript.htm"-->
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
						<%
							Dim oPara
							Const c_TITLE = 2
							Const c_SUBTITLE = 3
							Const c_SOURCE = 0
							Const c_DATE = 1
							Const c_AUTHOR = 4
							Const c_IMAGE = 5
							Const c_BODY = 6
							
							
							Set oNodeList = objXML.childNodes(1).childNodes
							
							s = "<font face=""verdana,arial,sans-serif"" color=""#6D1746"">"							
							s = s & "<h3>" & oNodeList(c_TITLE).text & "</h3>"
							s = s & "</font>"
							s = s & "<font face=""verdana,arial,sans-serif"" size=""2"" color=""#000000"">"	
							s = s & "<i>"
							
							If oNodeList(c_AUTHOR).text <> "" Then
								s = s &  "by " & oNodeList(c_AUTHOR).text 
							End If
							
							s = s & "&nbsp;from " & oNodeList(c_SOURCE).text 
							
							s = s &  "<BR>" & oNodeList(c_DATE).text
							
							s = s & "</i>"
							
							If oNodeList(c_SUBTITLE).text <> "" Then
								s = s & "<p><b>" & oNodeList(c_SUBTITLE).text & "</b></p>"
							End If
							
							For each oPara in oNodeList(c_BODY).childNodes
								s = s & "<p>" & oPara.text & "</p>"
							Next
							s = s & "</font>"
							Response.Write s
							
							Set oPara = Nothing
							Set oNode = Nothing
							Set oNodeList = Nothing
							Set objXML = Nothing
						%>
			
		</td>
	</tr>
	<tr>
		<td><font face="verdana,arial,sans-serif" size="2" color="#000000"><a href="news.asp">Back to News</a></font></td><td align="right"><img src="images/2bees.gif"></td>
	</tr>
	</table>
</div>
</body>
</html>
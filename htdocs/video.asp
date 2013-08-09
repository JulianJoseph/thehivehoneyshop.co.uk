<html>
<head>
	<title>The Hive Honey Shop : Inside The Hive</title>
	<!--#include file="include/javascript.htm"-->
    
<%@ Language="VBScript" %> 
<% Response.Expires= -1 
Response.AddHeader "Cache-Control", "no-cache" 
Response.AddHeader "Pragma", "no-cache" %> 
<% 
if Session("ct") = "" then 
fp = Server.MapPath("count.txt") 
Set fs = CreateObject("Scripting.FileSystemObject") 
Set a = fs.OpenTextFile(fp) 
ct = Clng(a.ReadLine) 
ct = ct + 1 
Session("ct") = ct 
a.close 
Set a = fs.CreateTextFile(fp, True) 
a.WriteLine(ct) 
a.Close 
Set a = Nothing 
Set fs = Nothing 
else 
ct = Clng(Session("ct")) 
end if 
%> 

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
		<td width="380">
			<font face="verdana,arial,sans-serif" color="#6D1746">
			<br>
			<h3> The Hive Honey Shop Video</h3>
			</font>		</td>
	</tr>
	<tr>
		<td colspan="2" align="center" valign="top"><font face="verdana,arial,sans-serif" size="2" color="#000000">&nbsp;			
		</font>
		  <table width="450" height="270" border="0" background="images/back.jpg">
            <tr>
              <td><object data="http://www.thehivehoneyshop.co.uk/images/flvplayer.swf?file=http://www.thehivehoneyshop.co.uk/images/beehive.flv&autostart=true&allowfullscreen=true" type="application/x-shockwave-flash" width="445" height="270" align="baseline" id="mpl">
                  <param name="movie" value="http://www.thehivehoneyshop.co.uk/images/flvplayer.swf?file=http://www.thehivehoneyshop.co.uk/images/beehive.flv&autostart=true&allowfullscreen=true" />
                  <param name="quality" value="high" />
                  <param name="flashvars"          value="enablejs=true&javascriptid=mpl" />
                </object></td>
            </tr>
          </table>
		  
		  <font face="verdana,arial,sans-serif" color="#6D1746"><h4><a href="http://www.thehivehoneyshop.co.uk">Click here to return to the Hive Honey Shop</a></h4>
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
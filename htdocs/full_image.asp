<html>
<head>
	<title>The Hive Honey Shop - Product Image</title>
	<script language=javascript>
	<!--
	function FitPic() { 
       iWidth = document.body.clientWidth; 
       iHeight = document.body.clientHeight; 
       iWidth = document.images[0].width - iWidth; 
       iHeight = document.images[0].height - iHeight; 
       window.resizeBy(iWidth, iHeight); 
       self.focus(); 
     }; 
	//-->
	</script>

</head>
<body bgcolor="#FFFFCC" onload="FitPic();" topmargin="0"  marginheight="0" leftmargin="0" marginwidth="0">
	<script language='javascript'> 
		var arrTemp=self.location.href.split("?"); 
		var picUrl = (arrTemp.length>0)?arrTemp[1]:""; 
		document.write( "<img src='" + picUrl + "' border=0>" ); 
	</script> 

</body>
</html>
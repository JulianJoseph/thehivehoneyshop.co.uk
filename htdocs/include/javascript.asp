<%
	Function getMenuItems
		Dim rs
		Dim sSQL
		Dim varArray()
		Dim i
		Dim Q
		If IsEmpty(Application("MenuItems"))  Then
			'Response.Write "//Read Database"
			Set rs = Server.CreateObject("ADODB.Recordset")
			sSQL = "SELECT Description, CategoryID FROM Category ORDER BY Description;"
			rs.Open sSQL, gsConnStr
			i = 0
			Q = Chr(34)
			Do While Not rs.EOF
				ReDim Preserve varArray(1, i)
				varArray(0,i) = rs("Description")
				varArray(1,i) = rs("CategoryID")
				i = i + 1
				rs.MoveNext 
			Loop	
			rs.Close
			Set rs = Nothing
			Application("MenuItems") = varArray
		End If
		getMenuItems = Application("MenuItems")
	End Function

%>
<script language="JavaScript">
	<!-- hide
	/* Function that displays status bar messages. */
	function MM_displayStatusMsg(msgStr)  { //v3.0
		status=msgStr; document.MM_returnValue = true;
	}
	function resetX()
	{
		x = (document.body.offsetWidth/2) + 150;	
	}

	function fwLoadMenus() {
	  if (window.fw_menu_0) return;
	  window.fw_menu_0 = new Menu("root",172,17,"Verdana, Arial, Helvetica, sans-serif",10,"#ffffcc","#ffffff","#6d1746","#000084");		
			<%
				Dim iMenu
				Dim menu
				menu = getMenuItems()
				For iMenu = 0 to Ubound(menu,2)
					'fw_menu_0.addMenuItem("Honeys","location='products.asp?category=1'");
					'response.write "//dynamic menu item" & vbcrlf
			%>
					fw_menu_0.addMenuItem("<%=menu(0,iMenu)%>","location='products.asp?category=<%=menu(1,iMenu)%>'");				
			<%Next%>
			fw_menu_0.bgImageUp="images/fwmenu1_172x17_up.gif";
			fw_menu_0.bgImageOver="images/fwmenu1_172x17_over.gif";
			fw_menu_0.hideOnMouseOut=true;
			fw_menu_0.writeMenus();
	} // fwLoadMenus()

	// stop hiding -->
</script>
<script language="JavaScript" src="fw_menu.js"></script>

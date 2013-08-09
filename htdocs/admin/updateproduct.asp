<!--#include file="adminutils.asp"-->
<%
	Response.Write request.Form 	
	
	Response.Buffer = true 
	Dim sSQL
	Dim Q
	Dim oConn
	Dim sConnStr
	Dim i
	Dim ProdID
	Dim rs
	
	Q = Chr(39)
	
	for i = 1 to request.Form.Count - 1 
		response.Write "<P>" & request.Form.Key(i)
		if LCase(Left(Request.Form.Key(i),9)) = "chkdelete" Then
			ProdID = Mid(Request.Form.Key(i), 10, len(request.Form.Key(i))-9)
			sSQL = "DELETE * FROM ProductDetail WHERE ProductID = " & ProdID
			Set oConn = Server.CreateObject("ADODB.Connection")
			oConn.Open gsConnStr
			oConn.Execute sSQL
			sSQL = "DELETE * FROM Products WHERE ProductID = " & ProdID
			oConn.Execute sSQL
			oConn.Close 
		end if
	next

	if request.Form("producttitle") <> "" then
		sSQL = "INSERT INTO Products(Title, CategoryID) VALUES(" & Q & request.Form("producttitle") & Q & ", " & Request.QueryString("categoryid") & ")"
		response.Write "<P>" & sSQL
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open gsConnStr
		oConn.Execute sSQL
		sSQL = "SELECT MAX(ProductID) FROM Products"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sSQL, oConn
		IF not rs.BOF and not rs.eof then
			ProdID = rs(0)
		end if
		rs.Close
		set rs = nothing 
		sSQL = "INSERT INTO ProductDetail(ProductID) VALUES(" & ProdID & ")"
		oConn.Execute sSQL
		response.Write "<P>" & sSQL
		oConn.Close 
	end if

	set oConn = nothing
	
	response.Redirect "list_products.asp?categoryid=" & Request.QueryString("categoryid")
%>
<%
	'Connection String
	gsConnStr="DSN=hivehoneyshop_data"
	
	'ADO Constants
	'---- CursorTypeEnum Values ----
	Const adOpenForwardOnly = 0
	Const adOpenKeyset = 1
	Const adOpenDynamic = 2
	Const adOpenStatic = 3


	'---- LockTypeEnum Values ----
	Const adLockReadOnly = 1
	Const adLockPessimistic = 2
	Const adLockOptimistic = 3
	Const adLockBatchOptimistic = 4
	
	'=================================================
	Function FixSpecialChars(s) 
		On Error Resume Next
	    If s <> "" Then
	        's = Replace(s, "&", "&amp;")
	        's = Replace(s, "'", "&apos;")
	        s = Replace(s, Chr(39), "''")
	        's = Replace(s, "<", "&lt;")
	        's = Replace(s, ">", "&gt;")
	        FixSpecialChars = s
	    Else
	        FixSpecialChars = ""
	    End If
	End Function
	'=================================================	
	Function FormatPara(s)
		On Error Resume Next
	    Dim aParas
	    Dim p
	    Dim i
	    aParas = Split(s, vbCrLf)
	    For i = 0 To UBound(aParas)
	        p = p & "<para>" & aParas(i) & "</para>"
	    Next
	    FormatPara = p
	End Function	
	'=================================================	

%>
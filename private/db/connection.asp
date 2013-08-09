<%
	Dim pwAdmin
	Dim sUserName
	Dim gsConnStr

	pwAdmin = "pr0p0l15"
	sUserName = "hamill"
	'gsConnStr= "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../../private/clients/hivehoneyshop/hive_data.mdb")
	gsConnStr= "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../private/db/hive_data.mdb")
	'gsConnStr="DSN=hivehoneyshop_data"

%>
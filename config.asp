<%	
	Const AdOpenKeySet=1
	Const AdLockOptimistic=3
	
	set banco=server.createobject("ADODB.Connection")
		'banco.open "PROVIDER=Microsoft.jet.OLEDB.4.0;Data Source=\\localhost\compos\composweb.mdb"
		banco.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\COMPOS\composweb.mdb"
		
	'strDBPath = "composweb.mdb"
		
	strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\COMPOS\composweb.mdb"
			
	AuthorCount = "SELECT Author.AuthorNick, Author.AuthorCount FROM Author ORDER BY Author.AuthorCount DESC"
	
	registros = 10
	relatorio = 30
	anoatual = 2007

	Set objConnConfig = Server.CreateObject("ADODB.Connection")
	Set objRsConfig = Server.CreateObject("ADODB.Recordset")
			
		objConnConfig.Open strConn
			
	objRsConfig.Open "SELECT TOP 1 * FROM Author WHERE AuthorLevel = 1", objConnConfig, 0, 1

		Application.Lock

			Application(ScriptName & "DefaultAuthorID") = objRsConfig("IDAuthor")

		Application.UnLock

	objRsConfig.Close		
	Set objRsConfig = Nothing
%>
<!--#include file="config.asp"-->
<%
	codsolic = request.querystring("nobreakpt_codigo")
		
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from nobreakpt WHERE nobreakpt_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic
				
	rsbanco.delete

	Response.Redirect("adicionar9.asp")
%>
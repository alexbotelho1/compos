<!--#include file="config.asp"-->
<%
	codsolic = request.querystring("config_codigo")
		
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from config WHERE config_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic
				
	rsbanco.delete

	Response.Redirect("adicionar9.asp")
%>
<!--#include file="config.asp"-->
<%
	codigo = request.querystring("estabilizador_codigo")
		
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from estabilizador WHERE estabilizador_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
				
	rsbanco.delete

	Response.Redirect("consultas4.asp")
%>
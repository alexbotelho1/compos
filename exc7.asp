<!--#include file="config.asp"-->
<%
	codigo = request.querystring("switch_codigo")
		
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from switch WHERE switch_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
				
	rsbanco.delete

	Response.Redirect("consultas4.asp")
%>
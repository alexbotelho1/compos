<!--#include file="config.asp"-->
<%
	codigo = request.querystring("computador_codigo")
		
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from computador WHERE computador_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
				
	rsbanco.delete

	Response.Redirect("consultas4.asp")
%>
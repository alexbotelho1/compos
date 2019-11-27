<!--#include file="config.asp"-->
<%
	codigo = request.querystring("nobreak_codigo")
		
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from nobreak WHERE nobreak_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
				
	rsbanco.delete

	Response.Redirect("consultas4.asp")
%>
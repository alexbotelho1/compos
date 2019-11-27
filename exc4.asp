<!--#include file="config.asp"-->
<%
	codigo = request.querystring("impressora_codigo")
		
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from impressora WHERE impressora_codigo = "&codigo&"",banco,AdOpenKeySet,AdLockOptimistic
				
	rsbanco.delete

	Response.Redirect("consultas4.asp")
%>
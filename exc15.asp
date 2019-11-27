<!--#include file="config.asp"-->
<%
	codsolic = request.querystring("secao_codigo")
		
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from secao WHERE secao_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic
				
	rsbanco.delete

	Response.Redirect("adicionar9.asp")
%>
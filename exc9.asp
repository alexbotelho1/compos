<!--#include file="config.asp"-->
<%
	codsolic = request.querystring("esquadrao_codigo")
		
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from esquadrao WHERE esquadrao_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic
				
	rsbanco.delete

	Response.Redirect("adicionar9.asp")
%>
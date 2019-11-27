<!--#include file="config.asp"-->
<%
	codsolic = request.querystring("marcaimp_codigo")
		
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from marcaimp WHERE marcaimp_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic
				
	rsbanco.delete

	Response.Redirect("adicionar9.asp")
%>
<!--#include file="config.asp"-->
<%
	codsolic = request.querystring("periferico_codigo")
		
	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from periferico WHERE periferico_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic
				
	rsbanco.delete

	Response.Redirect("adicionar9.asp")
%>
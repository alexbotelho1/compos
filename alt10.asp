<!--#include file="config.asp"-->
<%	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from config",banco,AdOpenKeySet,AdLockOptimistic

	config_manu=request.querystring("config_manu")
	
	altera = "Update config set config_manu='"&config_manu&"'"
	alterar = banco.execute(altera)
	
	Response.Redirect("adicionar9.asp") %>
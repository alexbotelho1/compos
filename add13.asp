<!--#include file="styles.asp"-->
<!--#include file="config.asp"-->
<% 	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from hardware",banco,AdOpenKeySet,AdLockOptimistic
		
If Trim(Request.querystring("hardware_nome")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else

	hardware_nome=request.querystring("hardware_nome")
	
	rsbanco.AddNew
	rsbanco("hardware_nome")=hardware_nome
	rsbanco.Update
	
End If	

	rsbanco.Close		
	banco.Close
	Set rsbanco = Nothing
	Set banco = Nothing
	
	Response.Redirect("adicionar9.asp") %>
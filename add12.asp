<!--#include file="styles.asp"-->
<!--#include file="config.asp"-->
<% 	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from esquadrao",banco,AdOpenKeySet,AdLockOptimistic
		
If Trim(Request.querystring("esquadrao_nome")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Voc� esqueceu de preencher um ou mais campos do formul�rio.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o bot�o de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else

	esquadrao_nome=request.querystring("esquadrao_nome")
	
	rsbanco.AddNew
	rsbanco("esquadrao_nome")=esquadrao_nome
	rsbanco.Update
	
End If	

	rsbanco.Close		
	banco.Close
	Set rsbanco = Nothing
	Set banco = Nothing
	
	Response.Redirect("adicionar9.asp") %>
<!--#include file="styles.asp"-->
<!--#include file="config.asp"-->
<% 	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from sti",banco,AdOpenKeySet,AdLockOptimistic
		
If Trim(Request.querystring("sti_nomeguerra")) = "" Then

		Response.Write("<br><br><p align='center'><font color='#FF0000' size='3'>Você esqueceu de preencher um ou mais campos do formulário.</p>")
		Response.Write("<br><br><p align='center'><font color='#ffffff' size='3'>Use o botão de retornar do navegador para corrigir o erro ou <a href='javascript:history.go(-1)'><font color='#FFFF00' size='3'>clique aqui</a>!</p>")
						
Else
	rsbanco.MoveLast
	sti_antiguidade = rsbanco("sti_antiguidade")
	sti_antiguidade = sti_antiguidade + 1
	
	sti_nomeguerra = request.querystring("sti_nomeguerra")
	
	rsbanco.AddNew
	rsbanco("sti_nomeguerra")=sti_nomeguerra
	rsbanco("sti_antiguidade")=sti_antiguidade
	rsbanco.Update
	
End If	

	rsbanco.Close		
	banco.Close
	Set rsbanco = Nothing
	Set banco = Nothing
	
	Response.Redirect("adicionar9.asp") %>
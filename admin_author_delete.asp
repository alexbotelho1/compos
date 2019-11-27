<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then 
	If Request.QueryString("mode") = "doit" Then
		Set objConn = Server.CreateObject("ADODB.Connection")
		Set objRs = Server.CreateObject("ADODB.Recordset")
		Set objRs2 = Server.CreateObject("ADODB.Recordset")
			objConn.Open strConn
				objRs.Open "SELECT * FROM Author WHERE IDAuthor = " & Request.QueryString("IDAuthor"), objConn, 3, 3		
					objRs.Delete 
				objRs.Close				
			objConn.Close 
		Set objConn = Nothing
		Set objRs = Nothing
		Set objRs2 = Nothing
			Application.Lock 			
				Application(ScriptName & "ConfigLoaded") = ""				
			Application.UnLock				
				Response.Write("<br><br><p align='center'><font color='#ffffff'>O Autor foi deletado com sucesso. <a href='admin_author.asp'>Click to continue</a>!</font></p>")
				Response.Write("<meta http-equiv='Refresh' content='1; URL=admin_author.asp'>")	
	Else
		Set objConn = Server.CreateObject("ADODB.Connection")
		Set objRs = Server.CreateObject("ADODB.Recordset")
			objConn.Open strConn
				objRs.Open "SELECT * FROM Author WHERE IDAuthor = " & Request.QueryString("IDAuthor"), objConn, 0, 1		
%>
<form action="admin_author_delete.asp?mode=doit&IDAuthor=<% = Request.QueryString("IDAuthor") %>" method="post">
<table border="1" width="600" height="102" align="center">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="504" height="102" align="center"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></td>
    </tr>
</table>			
	<table width="600" border="1" align="center">
        <tr>
			<td>			    
	            <table border="0" width="100%" cellspacing="1" cellpadding="4" bgcolor="#FFFFFF">
    		        <tr>
                		<td colspan="2"><b>Apagar Autores</b></td>
              		</tr>
              		<tr> 
                		<td colspan="2" valign="top"><b>Dados:</b></td>
              		</tr>
              		<tr> 
                		<td valign="top" align="right"> Login: </td>
                		<td valign="top"><input ReadOnly type="text" name="nickname" size="50" maxlength="50" value="<% = objRs("AuthorNick") %>"></td>
              		</tr>
              		<tr> 
                		<td valign="top" align="right"> Esquadrão: </td>
                		<td valign="top"><input ReadOnly type="text" name="esquadrao" size="50" maxlength="50" value="<% = objRs("AuthorEsquadrao") %>"></td>
              		</tr>              		
              		<tr>
                		<td valign="top" align="right"> Nome: </td>
                		<td valign="top"><input ReadOnly type="text" name="realname" size="50" maxlength="50" value="<% = objRs("AuthorRealName") %>"></td>
              		</tr>
              		<tr> 
                		<td valign="top" align="right"> Senha: </td>
                		<td valign="top"><input ReadOnly type="text" name="password" size="10" maxlength="10" value="<% = objRs("AuthorPassword") %>"></td>
              		</tr>
              		<tr>
                		<td valign="top" align="right"> N&iacute;vel<b></b>: </td>
                		<td valign="top">
                  			<select name="level">
                    			<option value="1"<% If objRs("AuthorLevel") = 1 Then Response.Write (" selected") %>>Administrator</option>
                    			<option value="2"<% If objRs("AuthorLevel") = 2 Then Response.Write (" selected") %>>Poder Adicionar/Alterar/Deletar</option>
                    			<option value="3"<% If objRs("AuthorLevel") = 3 Then Response.Write (" selected") %>>Poder Adicionar/Alterar</option>
								<option value="4"<% If objRs("AuthorLevel") = 4 Then Response.Write (" selected") %>>Cadastro de Hardware</option>                    			
                  			</select>                		
                		</td>
              		</tr>
              		<tr>
               			<td valign="top" align="right"> E-mail: </td>
                		<td valign="top"><input ReadOnly type="text" name="email" size="50" maxlength="100" value="<% = objRs("AuthorEmail") %>"></td>
              		</tr>
              		<tr> 
                		<td valign="top" align="right"> Contador: </td>
                		<td valign="top"><input type="text" name="count" size="3" maxlength="3" value="0" value="<% = objRs("AuthorCount") %>" ReadOnly></td>
              		</tr>
              		<tr> 
                		<td valign="top" align="center" colspan="2"><input type="submit" value=" D E L E T A R"></td>
              		</tr>
              		<tr> 
                		<td align="center" colspan="2"> <a class="Head" href="admin_author.asp">Voltar</a></td>
              		</tr>
            	</table>
			</td>
		</tr>
	</table>
<center><!--#include file="rodape.asp"--></center>
</form>
<%	
				objRs.Close
			objConn.Close 
		Set objConn = Nothing
		Set objRs = Nothing	
	End if
%>
<%
Else
	Response.Redirect("admin.asp")
End If
%>
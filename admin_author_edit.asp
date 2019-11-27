<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% Function ChkEmail(strTemp)

	ChkEmail = True
	strEmail = Trim(strTemp)	

	If Len(strEmail) > 0 Then

		intAtPos = InStr(1, strEmail, "@")

		If Not (intAtPos > 1 And (InStrRev(strEmail, ".") > intAtPos + 1)) Then
		
		    ChkEmail = False
		    
		ElseIf InStr(intAtPos + 1, strEmail, "@") > intAtPos Then
		
		    ChkEmail = False
		    
		ElseIf Mid(strEmail, intAtPos + 1, 1) = "." Then
		
		    ChkEmail = False
		    
		ElseIf InStr(1, Right(strEmail, 2), ".") > 0 Then
		
		    ChkEmail = False
		    
		End If
		
	End If

End Function
 If Session("LoggedIn") = True Then 		
	If Request.QueryString("mode") = "doit" Then
		If Trim(Request.Form("nickname")) = "" Or Trim(Request.Form("password")) = "" Or Trim(Request.Form("level")) = "" Then	
			Response.Write("<br><br><p align='center'><font color='#ffffff'>Você não preencheu um ou mais requisitos. Use o botão para voltar e corrigir.<a href='javascript:history.go(-1)'> Clique aqui!</a></font></p>")
		Else
			If Not ChkEmail(Trim(Request.Form("email"))) = True Then				
				Response.Write("<br><br><p align='center'><font color='#ffffff'>Você entrou com um endereço de e-mail errado. Use o botão para voltar e corrigir.<a href='javascript:history.go(-1)'> Clique aqui!</a></font></p>")
			Else
				Set objConn = Server.CreateObject("ADODB.Connection")
				Set objRs = Server.CreateObject("ADODB.Recordset")
				objConn.Open strConn
					objRs.Open "SELECT * FROM Author WHERE NOT IDAuthor = " & Request.QueryString("IDAuthor"), objConn, 0, 1		
						AuthorNickExists = False					
						Do While Not objRs.EOF							
							If objRs("AuthorNick") = Trim(Request.Form("nickname")) Then									
								AuthorNickExists = True								
							End If								
							objRs.MoveNext 							
						Loop						
							If AuthorNickExists = True Then						
								Response.Write("<br><br><p align='center'><font color='#ffffff'>Você escolheu um login que já existe. Use o botão para voltar e corrigir.<a href='javascript:history.go(-1)'> Clique aqui!</a></font></p>")						
							Else						
					objRs.Close
					objRs.Open "SELECT * FROM Author WHERE IDAuthor = " & Request.QueryString("IDAuthor"), objConn, 3, 3									
						objRs("AuthorNick") = Trim(Request.Form("nickname"))
						objRs("AuthorEsquadrao") = Trim(Request.Form("esquadrao"))						
						objRs("AuthorRealName") = Trim(Request.Form("realname"))
						objRs("AuthorPassword") = Trim(Request.Form("password"))
						objRs("AuthorLevel") = CInt(Request.Form("level"))
						objRs("AuthorEmail") = Trim(Request.Form("email"))
						objRs.Update 
							Response.Write("<br><br><p align='center'><font color='#ffffff'>O autor foi editado com sucesso. <a href='admin_author.asp'> Clique aqui!</a></font></p>")
							Response.Write("<meta http-equiv='Refresh' content='1; URL=admin_author.asp'>")					
							End If					
					objRs.Close
				objConn.Close 
			Set objConn = Nothing
			Set objRs = Nothing
				Application.Lock 				
					Application(ScriptName & "ConfigLoaded") = ""				
				Application.UnLock 										
			End If				
		End If	
	Else
		Set objConn = Server.CreateObject("ADODB.Connection")
		Set objRs = Server.CreateObject("ADODB.Recordset")
		objConn.Open strConn
			objRs.Open "SELECT * FROM Author WHERE IDAuthor = " & Request.QueryString("IDAuthor"), objConn, 0, 1		
%>
<form action="admin_author_edit.asp?mode=doit&IDAuthor=<% = Request.QueryString("IDAuthor") %>" method="post">
<table border="1" width="600" height="102" align="center">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="504" height="102" align="center"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></td>
    </tr>
</table>			
	<table width="594" border="1" align="center">
        <tr>
			<td width="602">			    
            	<table border="0" width="589" cellspacing="1" cellpadding="4" bgcolor="#FFFFFF">
              		<tr> 
                		<td colspan="2" width="589"> <b>Editar Autores</b></td>
              		</tr>
              		<tr> 
                		<td colspan="2" valign="top" width="589"> <b>Dados:</b></td>
              		</tr>
              		<tr> 
                		<td colspan="2" valign="top" width="589" align="center"> <font color="#999999">(Os campos marcados com um <b><font color="#FF9900">*</font></b> s&atilde;o necess&aacute;rios! )</font> </td>
              		</tr>
              		<tr>
                		<td valign="top" align="right" width="143"> Login: </td>
                		<td valign="top" width="459"><input type="text" name="nickname" size="50" maxlength="50" value="<% = objRs("AuthorNick") %>"></td>
              		</tr>
              		<tr>
                		<td valign="top" align="right" width="143"> Esquadrão: </td>
                		<td valign="top" width="459">
        	  				<select size="1" name="esquadrao">
  				  				<option<% If objRs("AuthorEsquadrao") = "EC" Then Response.Write (" selected") %>>
                                EC</option><option<% If objRs("AuthorEsquadrao") = "GSB" Then Response.Write (" selected") %>>GSB</option><option<% If objRs("AuthorEsquadrao") = "EI" Then Response.Write (" selected") %>>EI</option><option<% If objRs("AuthorEsquadrao") = "EP" Then Response.Write (" selected") %>>EP</option><option<% If objRs("AuthorEsquadrao") = "ES" Then Response.Write (" selected") %>>ES</option><option<% If objRs("AuthorEsquadrao") = "EIE" Then Response.Write (" selected") %>>EIE</option><option<% If objRs("AuthorEsquadrao") = "PAPV" Then Response.Write (" selected") %>>PAPV</option><option<% If objRs("AuthorEsquadrao") = "BINFA" Then Response.Write (" selected") %>>BINFA</option><option<% If objRs("AuthorEsquadrao") = "2GAV3" Then Response.Write (" selected") %>>2GAV3</option><option<% If objRs("AuthorEsquadrao") = "SCOAM" Then Response.Write (" selected") %>>SCOAM</option><option<% If objRs("AuthorEsquadrao") = "DTCEA-PV" Then Response.Write (" selected") %>>DTCEA-PV</option>
  			  				</select>                 		
  			  			</td>
              		</tr>              		
              		<tr>
                		<td valign="top" align="right" width="143"> Nome: </td>
                		<td valign="top" width="459"><input type="text" name="realname" size="50" maxlength="50" value="<% = objRs("AuthorRealName") %>"></td>
              		</tr>
              		<tr>
                		<td valign="top" align="right" width="143"> Senha <b></b>: </td>
                		<td valign="top" width="459"><input type="text" name="password" size="10" maxlength="10" value="<% = objRs("AuthorPassword") %>"></td>
              		</tr>
              		<tr>
                		<td valign="top" align="right" width="143"> N&iacute;vel<b></b>: </td>
                		<td valign="top" width="459"> 
                  			<select name="level">
                    			<option value="1"<% If objRs("AuthorLevel") = 1 Then Response.Write (" selected") %>>Administrator</option>
                    			<option value="2"<% If objRs("AuthorLevel") = 2 Then Response.Write (" selected") %>>Poder Adicionar/Alterar/Deletar</option>
                    			<option value="3"<% If objRs("AuthorLevel") = 3 Then Response.Write (" selected") %>>Poder Adicionar/Alterar</option>
								<option value="4"<% If objRs("AuthorLevel") = 4 Then Response.Write (" selected") %>>Cadastro de Hardware</option>                    			
                  			</select>
                		</td>
              		</tr>
              		<tr>
                		<td valign="top" align="right" width="143"> E-Mail: </td>
                		<td valign="top" width="459"><input type="text" name="email" size="50" maxlength="100" value="<% = objRs("AuthorEmail") %>"></td>
              		</tr>
              		<tr> 
                		<td valign="top" align="right" width="143"> Contador: </td>
                		<td valign="top" width="459"><input type="text" name="count" size="3" maxlength="3" value="0" value="<% = objRs("AuthorCount") %>" ReadOnly></td>
              		</tr>
              		<tr> 
                		<td valign="top" align="center" colspan="2" width="611"><input type="submit" value="E D I T A R">&nbsp;</td>
              		</tr>
              		<tr> 
                		<td class="Head" align="center" colspan="2" width="611"><a class="Head" href="admin_author.asp">Voltar</a></td>
              		</tr>
				</table>
			</td>
		</tr>
	</table>
<center><!--#include file="rodape.asp"--></center>	
</form>
<%				objRs.Close
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
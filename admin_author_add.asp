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
			Response.Write("<br><br><p align='center'><font color='#ffffff'>Você não preencheu um ou mais requisitos. Use o botão para voltar e corrigir.<a href='javascript:history.go(-1)'> Clique aqui!</a>!</font></p>")
		Else
			If Not Trim(Request.Form("password")) = Trim(Request.Form("confirm_password")) Then			
				Response.Write("<br><br><p align='center'><font color='#ffffff'>Você entrou com o Password de Confirmação errado. Use o botão para voltar e corrigir.<a href='javascript:history.go(-1)'> Clique aqui!</a>!</font></p>")
			Else			
				If Not ChkEmail(Trim(Request.Form("email"))) = True Then					
					Response.Write("<br><br><p align='center'><font color='#ffffff'>Você entrou com um endereço de e-mail errado. Use o botão para voltar e corrigir.<a href='javascript:history.go(-1)'> Clique aqui!</a>!</font></p>")
				Else
					Set objConn = Server.CreateObject("ADODB.Connection")
					Set objRs = Server.CreateObject("ADODB.Recordset")
					objConn.Open strConn
					objRs.Open "SELECT * FROM Author", objConn, 3, 3						
					AuthorNickExists = False						
						Do While Not objRs.EOF							
							If objRs("AuthorNick") = Trim(Request.Form("nickname")) Then								
								AuthorNickExists = True							
							End If							
							objRs.MoveNext 							
						Loop						
						If AuthorNickExists = True Then						
							Response.Write("<br><br><p align='center'><font color='#ffffff'>Você escolheu no login já existente. Use o botão para voltar e corrigir.<a href='javascript:history.go(-1)'> Clique aqui!</a>!</font></p>")						
						Else		
							objRs.AddNew 
							objRs("AuthorNick") = Trim(Request.Form("nickname"))
							objRs("AuthorEsquadrao") = Trim(Request.Form("esquadrao"))							
							objRs("AuthorRealName") = Trim(Request.Form("realname"))
							objRs("AuthorPassword") = Trim(Request.Form("password"))
							objRs("AuthorLevel") = CInt(Request.Form("level"))
							objRs("AuthorEmail") = Trim(Request.Form("email"))
							objRs("AuthorCount") = 0
							objRs.Update 					
							Response.Write("<br><br><p align='center'><font color='#ffffff'>O autor foi adicionado com sucesso. <a href='admin_author.asp'>Click to continue</a>!</font></p>")
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
		End If	
	Else
%>
<form action="admin_author_add.asp?mode=doit" method="post">
<table border="1" width="600" height="102" align="center">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="504" height="102" align="center"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></td>
    </tr>
</table>			
	<table width="600" border="1" align="center">
        <tr>
			<td width="600">
			    <table border="0" width="589" cellspacing="1" cellpadding="4" bgcolor="#FFFFFF">
			    	<tr>			        
                		<td colspan="2" width="589"><b>Adicionar Usuários</b></td>	        			        
			    	</tr>
			      	<tr>			        
			      		<td colspan="2" valign="top" width="589"><b>Dados:</b></td>
			      	</tr>
			      	<tr>			        
                		<td colspan="2" valign="top" width="589" align="center"> <font color="#999999">(Os campos marcados com um <b><font color="#FF9900">*</font></b> s&atilde;o necess&aacute;rios! )</font></td>
			      	</tr>			      
				  	<tr>				    
		                <td valign="top" align="right" width="148"> Login<b><font color="#FF9900">*</font></b>:</td>
						<td valign="top" width="431"><input type="text" name="nickname" size="45" maxlength="50"></td>
			      	</tr>
				  	<tr>				    
		                <td valign="top" align="right" width="148"> Esquadrão<b><font color="#FF9900">*</font></b>:</td>
						<td valign="top" width="431">
              				<select size="1" name="esquadrao">
  				  				<option value="0" selected>Selecione</option>            
  				  				<option>2GAV3</option>
  				  				<option>BINFA</option> 
  				  				<option>DTCEA-PV</option>  				  				 				            
  				  				<option>EC</option>
  				  				<option>EI</option>
  				  				<option>EIE</option>   				
  				  				<option>EP</option>
  				  				<option>ES</option> 				
  				  				<option>GSB</option>
  				  				<option>PAPV</option>
  				  				<option>SCOAM</option>
  			  				</select>						
						</td>
			      	</tr>			      	
			      	<tr>
					    <td valign="top" align="right" width="148"> Nome: </td>
						<td valign="top" width="431"><input type="text" name="realname" size="45" maxlength="50"></td>
					</tr>
			      	<tr>					
		                <td valign="top" align="right" width="148"> Senha<b><font color="#FF9900">*</font></b>:</td>
						<td valign="top" width="431"><input type="password" name="password" size="10" maxlength="10"></td>
			      	</tr>
			      	<tr>					
		                <td valign="top" align="right" width="148"> Confirmar Senha<b><font color="#FF9900">*</font></b>:</td>
						<td valign="top" width="431"><input type="password" name="confirm_password" size="10" maxlength="10"></td>
			      	</tr>
			      	<tr>					
                		<td valign="top" align="right" width="148"> N&iacute;vel<b><font color="#FF9900">*</font></b>:</td>
						<td valign="top" width="431">
							<select name="level">
                    			<option value="1">Administrador</option>
	                    		<option value="2">Poder Adicionar/Alterar/Deletar</option>
    	                		<option value="3">Poder Adicionar/Alterar</option>
    	                		<option value="4">Cadastro de Hardware</option>
        			        </select>
						</td>
			      	</tr>
			      	<tr>			
						<td valign="top" align="right" width="148">E-mail:</td>
						<td valign="top" width="431"><input type="text" name="email" size="45" maxlength="100"></td>
			      	</tr>	      				      			      
			      	<tr>			
						<td valign="top" align="center" colspan="2" width="580"><input type="submit" value="A D I C I O N A R"></td>
				  	</tr>	
				  	<tr>				    
                		<td class="Head" align="center" colspan="2" width="580" align="center"> <a href="admin_author.asp">Voltar</a></td>
				  	</tr>				  			  		      
			    </table>
			</td>
		</tr>
	</table>
<center><!--#include file="rodape.asp"--></center>	
</form>
<%		
	End if
%>
<%
Else
	Response.Redirect("admin.asp")
End If
%>
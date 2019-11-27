<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Request.QueryString("Logout") = "1" Then

	Session("LoggedIn") = False
	Response.Write("<meta http-equiv='Refresh' content='0; URL=admin.asp'>")
	
End If

If Session("LogError") < 3 Then
	
		If Not Session("LoggedIn") = True Then		
		
			If Request.QueryString("Login") = "1" Then
				 
				 Set objRs = Server.CreateObject("ADODB.Recordset")
				 Set objConn = Server.CreateObject("ADODB.Connection")			     
			     objConn.Open strConn
			     objRs.Open "SELECT TOP 1 IDAuthor, AuthorEsquadrao, AuthorNick, AuthorPassword, AuthorRealName, AuthorLevel FROM Author WHERE AuthorNick = '" &_
			      replace(Request.Form("Nickname"),"'","''") & "' AND AuthorPassword = '" &_
			       replace(Request.Form("Password"),"'","''") &_
			        "'", objConn, 0, 1
			     
			     If objRs.BOF And objRs.EOF Then
			     
					Session("LoggedIn") = False
					Response.Write("<br><br><p align='center'><font color='#ffffff' size='2'>Login ou senha incorreta <a href='admin.asp'><font color='#ffffff' size='2'>Tente outra vez</a>!</p>")					
					Session("LogError") = Session("LogError") + 1
			     
			     Else
			     
					Session("IDAuthor") = objRs("IDAuthor")
					Session("Nome") = objRs("AuthorRealName")
					Session("Esquadrao") = objRs("AuthorEsquadrao")
					Session("Level") = objRs("AuthorLevel")
					Session("LoggedIn") = True
					Response.Write("<br><br><p align='center'><font color='#ffffff' size='2'>Bem Vindo ao Administrador. Agora... <a href='admin.asp'><font color='#ffffff' size='2'>carrega aqui</a> para entrares!</p>")
					Response.Write("<meta http-equiv='Refresh' content='1; URL=admin.asp'>")
			     
			     End If
			     
			     objRs.Close
			     objConn.Close 
			     Set objConn = Nothing
			     Set objRs = Nothing
		
			Else %>
<body><center>
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="504" height="102" align="center"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></td>
    </tr>
</table>
<form name="FrontPage_Form1" action="admin.asp?Login=1" method="post" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
	<table border="1" cellpadding="2" cellspacing="2" align="center">
		<tr>
       		<td class="fundo2" align="center" colspan="2"><b>Administradores</b></td>
  		</tr>
		<tr>
 			<td class="fundo1" align="right" valign="middle">Usuário: </td>
      		<td class="fundo3" valign="middle"><input type="text" name="Nickname" size="20"> </td>
     	</tr>
  		<tr>
			<td class="fundo1" align="right" valign="middle">Senha : </td>
  			<td class="fundo3" valign="middle"><input type="password" name="Password" size="20" maxlength="10"></td>
		</tr>
		<tr>
			<td class="fundo2" colspan="2" align="center" valign="middle"><input type="submit" value="E N T R A R"></td>
		</tr>
	</table>
</form>
<form action="index.asp">
<table border="1" cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td class="fundo2" colspan="2" align="center" valign="middle"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></td>
	</tr>
</table></form>
    		<% End If	
		Else 
If Session("Level") <> 4 Then %>
<center>					   		
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" height="455">
    <tr>
      <td width="600" height="455" background="background.jpg">
        <center>
        <table border="1" cellpadding="0" cellspacing="0" style="border-width:0; border-collapse: collapse" bordercolor="#111111" width="600" height="244">
          <tr>
            <td width="600" height="148" style="border-style: none; border-width: medium" colspan="2">
            <p align="center" style="margin-top: 0; margin-bottom: 0"><img border="0" src="logo.gif"></p>
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="500" height="47">
                <tr>
                  <td class="fundo4" width="500" height="47" style="border-style: solid; border-width: 1; " align="center" style="margin-top: 0; margin-bottom: 0"><font color="#FFFFFF" size="2"><b>SISTEMA DE CONTROLE DE ORDEM DE SERVIÇO</b></font><p align="center" style="margin-top: 0; margin-bottom: 0"><i><font size="2" color="#FF0000"><b>ADMINISTRAÇÃO</b></font></i><p align="center" style="margin-top: 0; margin-bottom: 0">
                  <b><font size="2" color="#FFFF00">Bem Vindo!!! <% = Session("Nome") %> do <% = Session("Esquadrao") %></font></b></td>
                </tr>
              </table>
              </center>
			</td>
		</tr>
		<tr>
			<form action="index.asp">
				<td <% If Session("Level") = 1 Then %>width="300"<% Else %>width="600"<% End If %> height="61" align="center" style="border-style: none; border-width: medium" <% If Session("Level") <> 1 Then %>colspan="2"<% End If %>>
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>
<% If Session("Level") = 1 Then %>          	
			<form action="admin_author.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;Administrar Usuário&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>
<% End If %>          	          	                 	
        </tr>		
		<tr>
			<form action="cad_sol.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Solicitação&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>
<% If Session("Level") < 3 Then %>           	
			<form action="cad_hard.asp">
<% End If %>			
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Hardware&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>           	                	
        </tr>
		<tr>
			<form action="cad_os.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;Ordem de Serviço&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>
			<form action="adicionar9.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Tabelas&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>            	              	
        </tr>
		<tr>
			<form action="relatorio_os.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Relatórios&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>          	
			<form action="adicionar10.asp"> 			
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Configurações&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>          	               	
        </tr>    
		<tr>
			<form action="pesquisa_os.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Pesquisas&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form> 
			<form action="auditoria.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Auditoria&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>           	                	
        </tr>	        
        </table>        
      </td>
    </tr>
</table>
<table border="1" width="600" height="23">
	<tr>
		<td class="fundo2" align="center"><a href="admin.asp?Logout=1"><strong>Sair do Administrador</strong></a></td>
	</tr>
</table>
<% Else %>
<center>					   		
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" height="455">
    <tr>
      <td width="600" height="455" background="background.jpg" valign="top">
        <center>
        <table border="1" cellpadding="0" cellspacing="0" style="border-width:0; border-collapse: collapse" bordercolor="#111111" width="600" height="72">
          <tr>
            <td width="600" height="148" style="border-style: none; border-width: medium" colspan="2">
            <p align="center" style="margin-top: 0; margin-bottom: 0"><img border="0" src="logo.gif"></p>
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="500" height="47">
                <tr>
                  <td class="fundo4" width="500" height="47" style="border-style: solid; border-width: 1; " align="center" style="margin-top: 0; margin-bottom: 0"><font color="#FFFFFF" size="2"><b>SISTEMA DE CONTROLE DE ORDEM DE SERVIÇO</b></font><p align="center" style="margin-top: 0; margin-bottom: 0"><i><font size="2" color="#FF0000"><b>ADMINISTRAÇÃO</b></font></i><p align="center" style="margin-top: 0; margin-bottom: 0">
                  <b><font size="2" color="#FFFF00">Bem Vindo!!! <% = Session("Nome") %> do <% = Session("Esquadrao") %></font></b></td>
                </tr>
              </table>
              </center>
			</td>
		</tr>
		<tr>         	
			<td width="600" height="28" align="center" style="border-style: none; border-width: medium"></td>        	               	
        </tr> 		
		<tr>         	
			<form action="adicionar4.asp">	
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;Cadastrar Computador" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>          	               	
			<form action="adicionar5.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;Cadastrar Impressora&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>           	                	
        </tr>
		<tr>         	
			<form action="adicionar6.asp">	
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;Cadastrar NoBreak&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>          	               	
			<form action="adicionar7.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="Cadastrar Estabilizador" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>           	                	
        </tr>
		<tr>         	
			<form action="adicionar8.asp">	
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Cadastrar Switch&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>          	               	
			<form action="consultas4.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;Consultas Hardware&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>             	                	
        </tr>                
		<tr>
			<form action="index.asp">
				<td width="600" height="61" align="center" style="border-style: none; border-width: medium" colspan="2">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo4" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>           	                	
        </tr>             
        </table>        
      </td>
    </tr>
</table>
<table border="1" width="600" height="23">
	<tr>
		<td class="fundo2" align="center"><a href="admin.asp?Logout=1"><strong>Sair do Administrador</strong></a></td>
	</tr>
</table>
<%		End If	
 	End If
Else
	Response.Write("<br><br><p align='center'>Você fez três tentativas e agora terá fechar a sessão para reiniciar!</p>")
	Response.Write("<p align='center'>Se você esqueceu a sua senha, <a href='admin_reminder.asp'>clique aqui</a>!")
End If	%>
<center><!--#include file="rodape.asp"--></center></body></html>
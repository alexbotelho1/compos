<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then %>
<body><center>					   		
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
			<form action="admin.asp">
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
<!--#include file="rodape.asp"--></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>
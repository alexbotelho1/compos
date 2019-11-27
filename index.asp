<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<body><center>
<% 	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from config",banco,AdOpenKeySet,AdLockOptimistic

		rsbanco.MoveLast

	If rsbanco("config_manu") > 0 then %>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" height="455">
    <tr>
      <td width="600" height="455" background="background.jpg" valign="top">
        <center>
        <table border="1" cellpadding="0" cellspacing="0" style="border-width:0; border-collapse: collapse" bordercolor="#111111" width="315" height="439">
          <tr>
            <td width="315" height="56" style="border-style: none; border-width: medium; margin-top:0; margin-bottom:0" valign="top"><p align="center" style="margin-top: 0; margin-bottom: 0"><img border="0" src="logo.gif"></p>
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="301" height="50">
                <tr>
                  <td class="fundo4" width="315" height="50" style="border-style: solid; border-width: 1; " align="center"><font color="#FFFFFF" size="2"><b>SISTEMA DE CONTROLE DE ORDEM DE SERVIÇO</b></font></td>
                </tr>
              </table>
              </center>
			</td>
		</tr>
		<tr>
			<form action="admin.asp">
				<td width="315" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="301" height="50">
                		<tr>                  
                  			<td class="fundo4" width="301" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Administração&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>                  	
        </tr>		
		<tr>		
			<td width="315" height="205" align="center" style="border-style: none; border-width: medium">
            	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="301" height="50">
                	<tr>                  
                  		<td class="fundo4" width="301" height="50" style="border-style: solid; border-width: 1; " align="center">&nbsp;<p>
                        <b><font color="#FFFF00" size="4">Estamos em Manutenção</font></b></p>
                        <p><b><font color="#FFFF00" size="4">Logo retornaremos</font></b></p>
                        <p>&nbsp;</td> 
                	</tr>
             	</table>
        	</td>       	 	
        </tr>        
        </table>
        </center>
      </td>
    </tr>
</table>	
<% Else %>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" height="455">
    <tr>
      <td width="600" height="455" background="background.jpg" valign="top">
        <center>
        <table border="1" cellpadding="0" cellspacing="0" style="border-width:0; border-collapse: collapse" bordercolor="#111111" width="315" height="269">
          <tr>
            <td width="315" height="173" style="border-style: none; border-width: medium; margin-top:0; margin-bottom:0" valign="top"><p align="center" style="margin-top: 0; margin-bottom: 0"><img border="0" src="logo.gif"></p>
              <center>
              <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="301" height="50">
                <tr>
                  <td class="fundo4" width="315" height="50" style="border-style: solid; border-width: 1; " align="center"><font color="#FFFFFF" size="2"><b>SISTEMA DE CONTROLE DE ORDEM DE SERVIÇO</b></font></td>
                </tr>
              </table>
              </center>
			</td>
		</tr>
		<tr>
			<form action="admin.asp">
				<td width="315" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="301" height="50">
                		<tr>                  
                  			<td class="fundo4" width="301" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Administração&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>                  	
        </tr>		
		<tr>
			<form action="adicionar.asp">
				<td width="315" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="301" height="50">
                		<tr>                  
                  			<td class="fundo4" width="301" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;Abrir Solicitação&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>                  	
        </tr>
		<tr>
			<form action="consultas_completa.asp">
				<td width="315" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="301" height="50">
                		<tr>                  
                  			<td class="fundo4" width="301" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="Consultar Solicitação" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>                  	
        </tr>
<!--
		<tr>
			<form action="consultasos.asp">
				<td width="315" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="301" height="50">
                		<tr>                  
                  			<td class="fundo4" width="301" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Consultar OS&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>                  	
        </tr>
-->
		<tr>
			<form action="admin.asp">
				<td width="315" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="301" height="50">
                		<tr>                  
                  			<td class="fundo4" width="301" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="Cadastrar Hardware&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>                  	
        </tr>        
        </table>
        </center>
      </td>
    </tr>
</table>
<% End If %>
<!--#include file="rodape.asp"-->
</center>
</body></html>
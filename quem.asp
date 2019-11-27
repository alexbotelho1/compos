<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="styles.asp"-->
<!--#include file="config.asp"-->
<% 	ip = Request.ServerVariables("REMOTE_ADDR")
	host = Request.ServerVariables("REMOTE_HOST")
	logon = Request.ServerVariables("LOGON_USER")
	serverorigem = Request.ServerVariables("SERVER_NAME")
	authpassword = Request.ServerVariables("AUTH_PASSWORD")
	authuser = Request.ServerVariables("AUTH_USER")
	remoteuser = Request.ServerVariables("REMOTE_USER")
	httpuseragent = Request.ServerVariables("HTTP_USER_AGENT")	
	serveruser = Request.ServerVariables("SERVER_USER")
	remoteaddr = Request.ServerVariables("REMOTE_ADDR")		
	
	os_ip=ip
	os_host=host
	os_logon=logon
	os_server=serverorigem
	os_authpassword=authpassword 
	os_authuser=authuser 
	os_remoteuser=remoteuser 
	os_httpuseragent=httpuseragent
	os_serveruser=serveruser 
	os_remoteaddr=remoteaddr 

If os_server = "10.116.24.1" then
	os_server = "BAPV"
End If
	
%>
<body><center>
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="510" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Solicitação de abertura de Ordem de Serviço</b></font>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#008000"><b>Solicitação Inserida com Sucesso!!!</b></font></td>
    </tr>
</table>
<table border="1" width="600" height="23">
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>REMOTE_ADDR</b></td>
        <td class="fundo3" width="200" height="23" align="center"><% = os_ip %></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>REMOTE_HOST</b></td>
        <td class="fundo3" width="200" height="23" align="center"><% = os_host %></font></td>
      </tr>      
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>LOGON_USER</b></td>
        <td class="fundo3" width="200" height="23" align="center"><% = os_logon %></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>SERVER_NAME</b></td>
        <td class="fundo3" width="200" height="23" align="center"><% = os_server %></font></td>
      </tr>  
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>AUTH_PASSWORD</b></td>
        <td class="fundo3" width="200" height="23" align="center"><% = os_authpassword %></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>AUTH_USER</b></td>
        <td class="fundo3" width="200" height="23" align="center"><% = os_authuser %></font></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>REMOTE_USER</b></td>
        <td class="fundo3" width="200" height="23" align="center"><% = os_remoteuser %></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>HTTP_USER_AGENT</b></td>
        <td class="fundo3" width="200" height="23" align="center"><% = os_httpuseragent %></font></td>
      </tr>
      <tr>
        <td class="fundo1" width="100" height="23" align="center"><b>REMOTE_USER</b></td>
        <td class="fundo3" width="200" height="23" align="center"><% = os_serveruser %></font></td>
        <td class="fundo1" width="100" height="23" align="center"><b>REMOTE_ADDR</b></td>
        <td class="fundo3" width="200" height="23" align="center"><% = os_remoteaddr %></font></td>
      </tr>      
    </table>
<!--#include file="rodape.asp"--></form>
</center></body></html>
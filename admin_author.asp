<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then %>
<%	
	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject("ADODB.Recordset")
	objConn.Open strConn
	objRs.Open "SELECT IDAuthor, AuthorNick, AuthorEmail, AuthorLevel FROM Author ORDER BY AuthorLevel ASC, AuthorNick ASC", objConn, 0, 1
%>
<body>
<table border="1" width="600" height="102" align="center">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="504" height="102" align="center"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></td>
    </tr>
</table>
<table width="600" border="1" align="center">
	<tr>
	    <td>
	    	<table border="0" width="588" cellspacing="0" cellpadding="4" bgcolor="#FFFFFF" style="border-collapse: collapse" bordercolor="#111111">
	      		<tr>
	          		<td colspan="3" width="588"><b>Autores</b></td>	        
	      		</tr>
	      		<tr>
	          		<td valign="top" width="185"> <b>Nick </b>&nbsp;<a href="admin_author_add.asp"><small>(Adicionar Autor)</small></a> </td>
	          		<td valign="top" width="227"> <b>E-mail</b> </td>	        
	          		<td valign="top" width="146"> <b>Opções:</b></td>
	      		</tr>	
<% 
	Do While Not objRs.EOF
%>
				<tr>
			  		<td valign="top" width="185"><b><% = objRs("AuthorNick") %>&nbsp;</b><small>(N&iacute;vel:<b><font color="#0066FF"><% = objRs("AuthorLevel") %></font></b>)</small></td>
			  		<td valign="top" width="227"><a href="mailto:<% = objRs("AuthorEmail") %>"><% = objRs("AuthorEmail") %></a></td>				  
			  		<td valign="top" width="146">
			  			<a href="admin_author_edit.asp?IDAuthor=<% = objRS("IDAuthor") %>"><img border="0" src="icon_edit.gif" border="0" alt="Editar"> Editar</a> 
                		&nbsp;&nbsp;&nbsp;
                		<a href="admin_author_delete.asp?IDAuthor=<% = objRS("IDAuthor") %>"><img border="0" src="icon_delete.gif" border="0" alt="Apagar"> Deletar</a>
                	</td>
				</tr>	      
<%
	objRs.MoveNext		
		Loop
%>
				<tr>
			  		<td align="center" colspan="3" width="588"><a class="Head" href="admin.asp">Voltar ao Menu Principal</a></td>
		  		</tr>	
	    	</table>
		</td>
	</tr>
</table>
<center><!--#include file="rodape.asp"--></center>
<%	
	objRs.Close
	objConn.Close 
	Set objConn = Nothing
	Set objRs = Nothing	
%><% Else
	Response.Redirect("admin.asp")
End If %>
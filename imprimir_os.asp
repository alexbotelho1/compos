<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% codsolic = request.querystring("codsolic")

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os WHERE os_codigo = "&codsolic&"",banco,AdOpenKeySet,AdLockOptimistic
%>
<body topmargin="0" leftmargin="0"><center>
<table border="2" width="600" height="581" bordercolor="#000000" style="border-style:solid; border-width:2; border-collapse: collapse" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr>
    	<td width="97" height="101" valign="middle" align="center"><img border="0" src="logo.gif" width="73" height="90"></td>
    	<td width="503" height="101" align="center" style="border-right: 1px solid #000000; border-top: 1px solid #000000"><p style="margin-top: 0; margin-bottom: 0"><img border="0" src="aeronautica2.jpg" width="60" height="50"></p>
        <p style="margin-top: 0; margin-bottom: 0"><b><font size="4">COMANDO DA AERONÁUTICA</font></b></p>
        <p style="margin-top: 0; margin-bottom: 0">BASE AÉREA DE PORTO VELHO</p>
        <p style="margin-top: 0; margin-bottom: 0"><i><font size="2">Seção de Tecnologia da Informação</font></i></td>
    </tr>
    <tr>
    	<td width="600" height="20" colspan="2" align="center" style="border-style: solid; border-width: 1">
        <b>Ordem de Serviço n°</b> <%=rsbanco("os_numero")%> <b>de</b> <%=rsbanco("os_dataaber")%></td>
    </tr>
<% If rsbanco("os_solicperiferico") = "Login/Senha" Then %>
        <tr>
    	<td width="600" height="101" colspan="2">
			<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" height="142">
    			<tr>
      				<td width="31" height="137" rowspan="4" valign="middle">
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">S</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">O</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">L</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">I</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">T</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">A</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">Ç</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">Ã</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">O</font></b>
        			</td>
      				<td width="169" height="24"><b>&nbsp;N° da Solicitação:</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsbanco("os_codigo")%></td>
      				<td width="400" height="24"><b>&nbsp;Data da Solicitação:</b> <%=rsbanco("os_solicdata")%></td>
    			</tr>
    			<tr>
      				<td width="569" height="23" colspan="2">
      					<table border="0" width="570" height="23">
    						<tr>
      							<td width="448" height="23"><b>Do(a):</b> <%=rsbanco("os_solicmilitar")%>
                                <b>/</b> <%=rsbanco("os_solicsecao")%> <b>/</b> <%=rsbanco("os_solicesquadrao")%></td>
      							<td width="112" height="23"><b>Ramal:</b> <%=rsbanco("os_solicramal")%></td>
    						</tr>
  						</table>
      				</td>
    			</tr>
    			<tr>
      				<td width="569" height="25" colspan="2"><b>&nbsp;Para:</b> Chefe da Seção de Tecnologia da Informação</td>
    			</tr>
    			<tr>
      				<td width="569" height="64" colspan="2" valign="top" align="justify">&nbsp;<b>A(O):</b> <%=rsbanco("os_solicperiferico")%>. 
                    <b>Descrição do problema:</b> <font color="#FF0000"><b>Campo bloqueado por motivo de conter informações restritas aos administradores. Obrigado pela comprreensão e volte sempre!</b></font></td>
    			</tr>
			</table>   	
    	</td>
    </tr>
<% Else %>    
    <tr>
    	<td width="600" height="101" colspan="2">
			<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" height="142">
    			<tr>
      				<td width="31" height="137" rowspan="4" valign="middle">
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">S</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">O</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">L</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">I</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">T</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">A</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">Ç</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">Ã</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">O</font></b>
        			</td>
      				<td width="169" height="24"><b>&nbsp;N° da Solicitação:</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsbanco("os_codigo")%></td>
      				<td width="400" height="24"><b>&nbsp;Data da Solicitação:</b> <%=rsbanco("os_solicdata")%></td>
    			</tr>
    			<tr>
      				<td width="569" height="23" colspan="2">
      					<table border="0" width="570" height="23">
    						<tr>
      							<td width="448" height="23"><b>Do(a):</b> <%=rsbanco("os_solicmilitar")%>
                                <b>/</b> <%=rsbanco("os_solicsecao")%> <b>/</b> <%=rsbanco("os_solicesquadrao")%></td>
      							<td width="112" height="23"><b>Ramal:</b> <%=rsbanco("os_solicramal")%></td>
    						</tr>
  						</table>
      				</td>
    			</tr>
    			<tr>
      				<td width="569" height="25" colspan="2"><b>&nbsp;Para:</b> Chefe da Seção de Tecnologia da Informação</td>
    			</tr>
    			<tr>
      				<td width="569" height="64" colspan="2" valign="top" align="justify">&nbsp;<b>A(O):</b> <%=rsbanco("os_solicperiferico")%>. 
                    <b>Descrição do problema:</b> <%=rsbanco("os_solicdescricao")%>.</td>
    			</tr>
			</table>   	
    	</td>
    </tr>
<% End If %>
    <tr>
    	<td width="600" height="118" colspan="2">
			<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" height="102">
    			<tr>
      				<td width="30" height="97" rowspan="3" valign="middle">
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>A</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>B</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>E</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>R</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>T</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>U</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>R</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>A</b></font></p>
        			</td>
      				<td width="570" height="1"><b>Data da Abertura:</b> <%=rsbanco("os_dataaber")%></td>
    			</tr>
    			<tr>
      				<td width="570" height="23">
      					<table border="0" width="570" height="23">
    						<tr>
      							<td width="387" height="23"><b>Militar da STI:</b> <%=rsbanco("os_militaraber")%></td>
      							<td width="173" height="23"><b>Ramal da STI:</b> <%=rsbanco("os_ramalaber")%></td>
    						</tr>
  						</table>
      				</td>
    			</tr>
    			<tr>
      				<td width="570" height="82" valign="top"><b>Descrição:</b> <%=rsbanco("os_descricaoaber")%></td>
    			</tr>
			</table>    	
    	</td>
    </tr>
    <tr>
    	<td width="600" height="107" colspan="2">
			<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" height="174">
    			<tr>
      				<td width="31" height="169" rowspan="4" valign="middle">
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">E</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">X</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">E</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">C</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">U</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">Ç</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">Ã</font></b></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">O</font></b>
        			</td>
      				<td width="169" height="1"><b>&nbsp;Homem/Hora:</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%=rsbanco("os_tempoexec")%></td>
      				<td width="400" height="1"><b>&nbsp;Data da Execução:</b> <%=rsbanco("os_dataexec")%></td>
    			</tr>
    			<tr>
      				<td width="569" height="1" colspan="2"><b>&nbsp;Militar que Executou:</b> <%=rsbanco("os_militarexec")%></td>
    			</tr>    			
    			<tr>
      				<td width="569" height="60" colspan="2">
      					<table border="0" width="570" height="136" style="border-collapse: collapse" bordercolor="#000000" cellpadding="0" cellspacing="0">
    						<tr>
      							<td width="251" height="136" valign="top" style="border-right-style:solid; border-right-width:1"><b>&nbsp;Serviço Executado:</b> <%=rsbanco("os_descricaoexec")%></td>
      							<td width="309" height="136" valign="top">&nbsp;<b>Material Utilizado:</b> <%=rsbanco("os_matusadoexec")%></td>
    						</tr>
  						</table>
      				</td>
    			</tr>
			</table>    	
    	</td>
    </tr>
    <tr>
    	<td width="600" height="117" colspan="2">
			<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="600" height="158">
    			<tr>
      				<td width="27" height="153" rowspan="4" valign="middle">
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>F</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>E</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>C</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>H</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>A</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>M</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>E</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>N</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2"><b>T</b></font></p>
        				<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="2">O</font></b>
        			</td>
      				<td width="257" height="19">&nbsp;<b>Militar que Fechou:</b> <%=rsbanco("os_militarconc")%></td>
      				<td width="316" height="19">&nbsp;<b>Data do Fechamento:</b> <%=rsbanco("os_dataconc")%></td>
    			</tr>
    			<tr>
      				<td width="573" height="7" colspan="2">&nbsp;<b>Militar que Recebeu:</b> <%=rsbanco("os_milrecconc")%></td>
    			</tr>    			
    			<tr>
      				<td width="252" height="85" colspan="2" valign="top">&nbsp;<b>Observação Final:</b> <%=rsbanco("os_observconc")%></td>
      			</tr>
      		<% If rsbanco("os_status") = 1 then %>
      			<tr>
      				<td width="311" height="4" colspan="2">&nbsp;<b>Status da Ordem de Serviço:</b> Aberta</td>
    			</tr>
    		<% Else
    			If rsbanco("os_status") = 2 then %>
      			<tr>
      				<td width="311" height="4" colspan="2">&nbsp;<b>Status da Ordem de Serviço:</b> Executada</td>
    			</tr>
    			<% Else 
    				If rsbanco("os_status") = 3 then %>
      			<tr>
      				<td width="311" height="4" colspan="2">&nbsp;<b>Status da Ordem de Serviço:</b> Fechada</td>
    			</tr>
    				<% End If 
    			End IF
    		End If %>
			</table>     	
    	</td>
    </tr>
</table></center></body></html>
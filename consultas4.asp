<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	If Request.QueryString("Ordem") = "" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from computador order by computador_esquadrao ASC, computador_secao ASC, computador_tipo ASC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco1=server.createobject("ADODB.Recordset")
			rsbanco1.open "Select * from impressora order by impressora_esquadrao ASC, impressora_secao ASC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco2=server.createobject("ADODB.Recordset")
			rsbanco2.open "Select * from nobreak order by nobreak_esquadrao ASC, nobreak_secao ASC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco3=server.createobject("ADODB.Recordset")
			rsbanco3.open "Select * from estabilizador order by estabilizador_secao ASC, estabilizador_esquadrao ASC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco4=server.createobject("ADODB.Recordset")
			rsbanco4.open "Select * from switch order by switch_esquadrao ASC, switch_secao ASC",banco,AdOpenKeySet,AdLockOptimistic												
	End If
	If Request.QueryString("Ordem") = "1" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from computador order by computador_fcg ASC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco1=server.createobject("ADODB.Recordset")
			rsbanco1.open "Select * from impressora",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco2=server.createobject("ADODB.Recordset")
			rsbanco2.open "Select * from nobreak",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco3=server.createobject("ADODB.Recordset")
			rsbanco3.open "Select * from estabilizador",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco4=server.createobject("ADODB.Recordset")
			rsbanco4.open "Select * from switch",banco,AdOpenKeySet,AdLockOptimistic												
	End If
	If Request.QueryString("Ordem") = "2" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from computador order by computador_fcg DESC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco1=server.createobject("ADODB.Recordset")
			rsbanco1.open "Select * from impressora",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco2=server.createobject("ADODB.Recordset")
			rsbanco2.open "Select * from nobreak",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco3=server.createobject("ADODB.Recordset")
			rsbanco3.open "Select * from estabilizador",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco4=server.createobject("ADODB.Recordset")
			rsbanco4.open "Select * from switch",banco,AdOpenKeySet,AdLockOptimistic												
	End If
	If Request.QueryString("Ordem") = "3" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from computador",banco,AdOpenKeySet,AdLockOptimistic	
		set rsbanco1=server.createobject("ADODB.Recordset")
			rsbanco1.open "Select * from impressora order by impressora_fcg ASC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco2=server.createobject("ADODB.Recordset")
			rsbanco2.open "Select * from nobreak",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco3=server.createobject("ADODB.Recordset")
			rsbanco3.open "Select * from estabilizador",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco4=server.createobject("ADODB.Recordset")
			rsbanco4.open "Select * from switch",banco,AdOpenKeySet,AdLockOptimistic										
	End If
	If Request.QueryString("Ordem") = "4" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from computador",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco1=server.createobject("ADODB.Recordset")
			rsbanco1.open "Select * from impressora order by impressora_fcg DESC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco2=server.createobject("ADODB.Recordset")
			rsbanco2.open "Select * from nobreak",banco,AdOpenKeySet,AdLockOptimistic		
		set rsbanco3=server.createobject("ADODB.Recordset")
			rsbanco3.open "Select * from estabilizador",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco4=server.createobject("ADODB.Recordset")
			rsbanco4.open "Select * from switch",banco,AdOpenKeySet,AdLockOptimistic										
	End If
	If Request.QueryString("Ordem") = "5" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from computador",banco,AdOpenKeySet,AdLockOptimistic			
		set rsbanco1=server.createobject("ADODB.Recordset")
			rsbanco1.open "Select * from impressora",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco2=server.createobject("ADODB.Recordset")
			rsbanco2.open "Select * from nobreak order by nobreak_fcg ASC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco3=server.createobject("ADODB.Recordset")
			rsbanco3.open "Select * from estabilizador",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco4=server.createobject("ADODB.Recordset")
			rsbanco4.open "Select * from switch",banco,AdOpenKeySet,AdLockOptimistic					
	End If
	If Request.QueryString("Ordem") = "6" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from computador",banco,AdOpenKeySet,AdLockOptimistic			
		set rsbanco1=server.createobject("ADODB.Recordset")
			rsbanco1.open "Select * from impressora",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco2=server.createobject("ADODB.Recordset")
			rsbanco2.open "Select * from nobreak order by nobreak_fcg DESC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco3=server.createobject("ADODB.Recordset")
			rsbanco3.open "Select * from estabilizador",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco4=server.createobject("ADODB.Recordset")
			rsbanco4.open "Select * from switch",banco,AdOpenKeySet,AdLockOptimistic			
	End If
	If Request.QueryString("Ordem") = "7" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from computador",banco,AdOpenKeySet,AdLockOptimistic			
		set rsbanco1=server.createobject("ADODB.Recordset")
			rsbanco1.open "Select * from impressora",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco2=server.createobject("ADODB.Recordset")
			rsbanco2.open "Select * from nobreak",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco3=server.createobject("ADODB.Recordset")
			rsbanco3.open "Select * from estabilizador order by estabilizador_fcg ASC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco4=server.createobject("ADODB.Recordset")
			rsbanco4.open "Select * from switch",banco,AdOpenKeySet,AdLockOptimistic						
	End If
	If Request.QueryString("Ordem") = "8" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from computador",banco,AdOpenKeySet,AdLockOptimistic			
		set rsbanco1=server.createobject("ADODB.Recordset")
			rsbanco1.open "Select * from impressora",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco2=server.createobject("ADODB.Recordset")
			rsbanco2.open "Select * from nobreak",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco3=server.createobject("ADODB.Recordset")
			rsbanco3.open "Select * from estabilizador order by estabilizador_fcg DESC",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco4=server.createobject("ADODB.Recordset")
			rsbanco4.open "Select * from switch",banco,AdOpenKeySet,AdLockOptimistic								
	End If
	If Request.QueryString("Ordem") = "9" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from computador",banco,AdOpenKeySet,AdLockOptimistic			
		set rsbanco1=server.createobject("ADODB.Recordset")
			rsbanco1.open "Select * from impressora",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco2=server.createobject("ADODB.Recordset")
			rsbanco2.open "Select * from nobreak",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco3=server.createobject("ADODB.Recordset")
			rsbanco3.open "Select * from estabilizador",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco4=server.createobject("ADODB.Recordset")
			rsbanco4.open "Select * from switch order by switch_fcg ASC",banco,AdOpenKeySet,AdLockOptimistic						
	End If
	If Request.QueryString("Ordem") = "10" Then
		set rsbanco=server.createobject("ADODB.Recordset")
			rsbanco.open "Select * from computador",banco,AdOpenKeySet,AdLockOptimistic			
		set rsbanco1=server.createobject("ADODB.Recordset")
			rsbanco1.open "Select * from impressora",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco2=server.createobject("ADODB.Recordset")
			rsbanco2.open "Select * from nobreak",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco3=server.createobject("ADODB.Recordset")
			rsbanco3.open "Select * from estabilizador",banco,AdOpenKeySet,AdLockOptimistic
		set rsbanco4=server.createobject("ADODB.Recordset")
			rsbanco4.open "Select * from switch order by switch_fcg DESC",banco,AdOpenKeySet,AdLockOptimistic								
	End If
	
	qtdcomp = 0
	If Session("Level") = 4 then
		Do While Not rsbanco.EOF
			If rsbanco("computador_esquadrao") = Session("Esquadrao") then
				qtdcomp = qtdcomp + 1
				rsbanco.Movenext
			Else
				rsbanco.MoveNext
			End If
		Loop
	Else
		Do While Not rsbanco.EOF
			qtdcomp = qtdcomp + 1
			rsbanco.movenext
		Loop
	End If
	If qtdcomp > 0 then
		rsbanco.moveFirst
	End If
	
	qtdimp = 0
	If Session("Level") = 4 then
		Do While Not rsbanco1.EOF
			If rsbanco1("computador_esquadrao") = Session("Esquadrao") then
				qtdimp = qtdimp + 1
				rsbanco1.movenext
			Else
				rsbanco1.MoveNext
			End If
		Loop
	Else
		Do While Not rsbanco1.EOF
			qtdimp = qtdimp + 1
			rsbanco1.movenext
		Loop
	End If
	If qtdimp > 0 then
		rsbanco1.moveFirst
	End If
	
	qtdnobreak = 0
	If Session("Level") = 4 then
		Do While Not rsbanco2.EOF
			If rsbanco2("computador_esquadrao") = Session("Esquadrao") then
				qtdnobreak = qtdnobreak + 1
				rsbanco2.movenext
			Else
				rsbanco2.MoveNext
			End If
		Loop
	Else
		Do While Not rsbanco2.EOF
			qtdnobreak = qtdnobreak + 1
			rsbanco2.movenext
		Loop
	End If
	If qtdnobreak > 0 then
		rsbanco2.moveFirst
	End If
	
	qtdestab = 0
	If Session("Level") = 4 then
		Do While Not rsbanco3.EOF
			If rsbanco3("computador_esquadrao") = Session("Esquadrao") then
				qtdestab = qtdestab + 1
				rsbanco3.movenext
			Else
				rsbanco3.MoveNext
			End If
		Loop
	Else
		Do While Not rsbanco3.EOF
			qtdestab = qtdestab + 1
			rsbanco3.movenext
		Loop
	End If
	If qtdestab > 0 then
		rsbanco3.moveFirst
	End If
	
	qtdswitch = 0
	If Session("Level") = 4 then
		Do While Not rsbanco4.EOF
			If rsbanco4("computador_esquadrao") = Session("Esquadrao") then
				qtdswitch = qtdswitch + 1
				rsbanco4.movenext
			Else
				rsbanco4.MoveNext
			End If
		Loop
	Else
		Do While Not rsbanco4.EOF
			qtdswitch = qtdswitch + 1
			rsbanco4.movenext
		Loop
	End If
	If qtdswitch > 0 then
		rsbanco4.moveFirst
	End If %>
<body><center>
<% If Session("Level") = 4 Then %>
<form action="admin.asp">
<% Else %>
<form action="cad_hard.asp">
<% End If %>
<table border="1" width="740" height="102">
    <tr>
      <td class="fundo1" width="100" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="640" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Consulta dos Cadastro de Hardware</b></font></td>
    </tr>
</table>
<table width="740" align="center" border="1">
<tr>
	  <td class="fundo1" width="740" height="29" align="center" colspan="10"><b><font color="#7F0D11" size="2"><% = qtdcomp %></font><font size="2"> Computador(es)</td>	  
</tr>
<tr class="fundo3">
	  <td width="58" height="29" align="center"><a href="consultas4.asp?Ordem=1"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;<font size="2"><b>FCG</b></font>&nbsp;<a href="consultas4.asp?Ordem=2"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
      <td width="112" height="29" align="center"><b><font size="2">Seção</td>
	  <td width="130" height="29" align="center"><b><font size="2">Esquadrão</td>      
      <td width="90" height="29" align="center"><b><font size="2">Tipo</td>
      <td width="140" height="29" align="center"><b><font size="2">Sist. Oper.</td>
      <td width="37" height="29" align="center"><b><font size="2">Qtd Proc</td>
      <td width="57" height="29" align="center"><b><font size="2">Freq</td>
      <td width="53" height="29" align="center"><b><font size="2">Mem</td>
	  <td width="56" height="29" align="center"><b><font size="2">HD</td>
	  <td width="57" height="29" align="center"><b><font size="2">Opções</td>	  
</tr>
<% If Session("Level") = 4 then
	Do While Not rsbanco.EOF
		If rsbanco("computador_esquadrao") = Session("Esquadrao") then %>
<tr class="fundo5">
	  <td width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td width="140" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td width="37" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %>&nbsp;<% = rsbanco("computador_procfreq") %></td>
      <td width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %>&nbsp;<% = rsbanco("computador_memocapa") %></td>
	  <td width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %>&nbsp;<% = rsbanco("computador_hdcapa") %></td>	
      <td width="57" height="29" align="center"><font size="2">
      <a href="alt_computador.asp?codigo=<% = rsbanco("computador_codigo") %>"><img border="0" src="editar.gif" alt="Editar Computador"></a>&nbsp;&nbsp;
      <% If Session("Level") < 3 Then %><a href="exc_computador.asp?codigo=<% = rsbanco("computador_codigo") %>"><img border="0" src="del.gif" alt="Excluir Computador"></a><% End If %></td>      
    </tr>
<%			rsbanco.movenext
		Else
			rsbanco.MoveNext
		End If
	Loop
Else
	Do While Not rsbanco.EOF %>
<tr class="fundo5">
	  <td width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td width="140" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td width="37" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %>&nbsp;<% = rsbanco("computador_procfreq") %></td>
      <td width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %>&nbsp;<% = rsbanco("computador_memocapa") %></td>
	  <td width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %>&nbsp;<% = rsbanco("computador_hdcapa") %></td>	
      <td width="57" height="29" align="center"><font size="2">
      <a href="alt_computador.asp?codigo=<% = rsbanco("computador_codigo") %>"><img border="0" src="editar.gif" alt="Editar Computador"></a>&nbsp;&nbsp;
      <% If Session("Level") < 3 Then %><a href="exc_computador.asp?codigo=<% = rsbanco("computador_codigo") %>"><img border="0" src="del.gif" alt="Excluir Computador"></a><% End If %></td>      
    </tr>
<%		rsbanco.movenext
	Loop
End If %>	
</table>
<!-- Impressora -->
<table width="740" align="center" border="1">
<tr>
	  <td class="fundo1" width="740" height="29" align="center" colspan="10"><b><font color="#7F0D11" size="2"><% = qtdimp %></font><font size="2"> Impressora(s)</td>	  
</tr>
<tr class="fundo3">
	  <td width="58" height="29" align="center"><a href="consultas4.asp?Ordem=3"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;<font size="2"><b>FCG</b></font>&nbsp;<a href="consultas4.asp?Ordem=4"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
      <td width="112" height="29" align="center"><b><font size="2">Seção</td>
	  <td width="130" height="29" align="center"><b><font size="2">Esquadrão</td>      
      <td width="90" height="29" align="center"><b><font size="2">Marca</td>
      <td width="100" height="29" align="center"><b><font size="2">Modelo</td>
      <td width="77" height="29" align="center"><b><font size="2">Impressão</td>
      <td width="57" height="29" align="center"><b><font size="2">Cor</td>
      <td width="53" height="29" align="center"><b><font size="2">Consumo</td>
	  <td width="56" height="29" align="center"><b><font size="2">Situação</td>
	  <td width="57" height="29" align="center"><b><font size="2">Opções</td>	  
</tr>
<% If Session("Level") = 4 then
	Do While Not rsbanco1.EOF
		If rsbanco1("impressora_esquadrao") = Session("Esquadrao") then %>
<tr class="fundo5">
	  <td width="58" height="29" align="center"><font size="2"><% = rsbanco1("impressora_fcg") %></td>
      <td width="112" height="29" align="center"><font size="2"><% = rsbanco1("impressora_secao") %></td>
      <td width="130" height="29" align="center"><font size="2"><% = rsbanco1("impressora_esquadrao") %></td>      
      <td width="90" height="29" align="center"><font size="2"><% = rsbanco1("impressora_marca") %></td>
      <td width="100" height="29" align="center"><font size="2"><% = rsbanco1("impressora_modelo") %></td>
      <td width="77" height="29" align="center"><font size="2"><% = rsbanco1("impressora_impressao") %></td>
      <td width="57" height="29" align="center"><font size="2"><% = rsbanco1("impressora_cor") %></td>
      <td width="53" height="29" align="center"><font size="2"><% = rsbanco1("impressora_consumo") %></td>
	  <td width="56" height="29" align="center"><font size="2"><% = rsbanco1("impressora_situacao") %></td>	
      <td width="57" height="29" align="center"><font size="2">
      <a href="alt_impressora.asp?codigo=<% = rsbanco1("impressora_codigo") %>"><img border="0" src="editar.gif" alt="Editar Impressora"></a>&nbsp;&nbsp;
      <% If Session("Level") < 3 Then %><a href="exc_impressora.asp?codigo=<% = rsbanco1("impressora_codigo") %>"><img border="0" src="del.gif" alt="Excluir Impressora"></a><% End If %></td>      
    </tr>
<%			rsbanco1.movenext
		Else
			rsbanco1.MoveNext
		End If
	Loop
Else
	Do While Not rsbanco1.EOF %>
<tr class="fundo5">
	  <td width="58" height="29" align="center"><font size="2"><% = rsbanco1("impressora_fcg") %></td>
      <td width="112" height="29" align="center"><font size="2"><% = rsbanco1("impressora_secao") %></td>
      <td width="130" height="29" align="center"><font size="2"><% = rsbanco1("impressora_esquadrao") %></td>      
      <td width="90" height="29" align="center"><font size="2"><% = rsbanco1("impressora_marca") %></td>
      <td width="100" height="29" align="center"><font size="2"><% = rsbanco1("impressora_modelo") %></td>
      <td width="77" height="29" align="center"><font size="2"><% = rsbanco1("impressora_impressao") %></td>
      <td width="57" height="29" align="center"><font size="2"><% = rsbanco1("impressora_cor") %></td>
      <td width="53" height="29" align="center"><font size="2"><% = rsbanco1("impressora_consumo") %></td>
	  <td width="56" height="29" align="center"><font size="2"><% = rsbanco1("impressora_situacao") %></td>	
      <td width="57" height="29" align="center"><font size="2">
      <a href="alt_impressora.asp?codigo=<% = rsbanco1("impressora_codigo") %>"><img border="0" src="editar.gif" alt="Editar Impressora"></a>&nbsp;&nbsp;
      <% If Session("Level") < 3 Then %><a href="exc_impressora.asp?codigo=<% = rsbanco1("impressora_codigo") %>"><img border="0" src="del.gif" alt="Excluir Impressora"></a><% End If %></td>      
    </tr>
<%		rsbanco1.movenext
	Loop
End If %>	
</table>
<!-- NoBreak -->
<table width="740" align="center" border="1">
<tr>
	  <td class="fundo1" width="740" height="29" align="center" colspan="8"><b><font color="#7F0D11" size="2"><% = qtdnobreak %></font><font size="2"> NoBreak(s)</td>	  
</tr>
<tr class="fundo3">
	  <td width="60" height="29" align="center"><a href="consultas4.asp?Ordem=5"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;<font size="2"><b>FCG</b></font>&nbsp;<a href="consultas4.asp?Ordem=6"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
      <td width="130" height="29" align="center"><b><font size="2">Seção</td>
	  <td width="140" height="29" align="center"><b><font size="2">Esquadrão</td>      
      <td width="90" height="29" align="center"><b><font size="2">Marca</td>
      <td width="110" height="29" align="center"><b><font size="2">Potência</td>
      <td width="90" height="29" align="center"><b><font size="2">Qtd Saída</td>
	  <td width="60" height="29" align="center"><b><font size="2">Situação</td>
	  <td width="60" height="29" align="center"><b><font size="2">Opções</td>	  
</tr>
<% If Session("Level") = 4 then
	Do While Not rsbanco2.EOF
		If rsbanco2("nobreak_esquadrao") = Session("Esquadrao") then %>
<tr class="fundo5">
	  <td width="60" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_fcg") %></td>
      <td width="130" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_secao") %></td>
      <td width="140" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_esquadrao") %></td>      
      <td width="90" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_marca") %></td>
      <td width="110" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_potencia") %> KVa</td>
      <td width="90" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_saida") %></td>
	  <td width="60" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_situacao") %></td>	
      <td width="60" height="29" align="center"><font size="2">
      <a href="alt_nobreak.asp?codigo=<% = rsbanco2("nobreak_codigo") %>"><img border="0" src="editar.gif" alt="Editar NoBreak"></a>&nbsp;&nbsp;
      <% If Session("Level") < 3 Then %><a href="exc_nobreak.asp?codigo=<% = rsbanco2("nobreak_codigo") %>"><img border="0" src="del.gif" alt="Excluir Nobreak"></a><% End If %></td>      
    </tr>
<%			rsbanco2.movenext
		Else
			rsbanco2.MoveNext
		End If
	Loop
Else
	Do While Not rsbanco2.EOF %>
<tr class="fundo5">
	  <td width="60" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_fcg") %></td>
      <td width="130" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_secao") %></td>
      <td width="140" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_esquadrao") %></td>      
      <td width="90" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_marca") %></td>
      <td width="110" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_potencia") %> KVa</td>
      <td width="90" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_saida") %></td>
	  <td width="60" height="29" align="center"><font size="2"><% = rsbanco2("nobreak_situacao") %></td>	
      <td width="60" height="29" align="center"><font size="2">
      <a href="alt_nobreak.asp?codigo=<% = rsbanco2("nobreak_codigo") %>"><img border="0" src="editar.gif" alt="Editar Nobreak"></a>&nbsp;&nbsp;
      <% If Session("Level") < 3 Then %><a href="exc_nobreak.asp?codigo=<% = rsbanco2("nobreak_codigo") %>"><img border="0" src="del.gif" alt="Excluir Nobreak"></a><% End If %></td>      
    </tr>
<%		rsbanco2.movenext
	Loop
End If %>	
</table>
<!-- Estabilizador -->
<table width="740" align="center" border="1">
<tr>
	  <td class="fundo1" width="740" height="29" align="center" colspan="8"><b><font color="#7F0D11" size="2"><% = qtdestab %></font><font size="2"> Estabilizador(es)</td>	  
</tr>
<tr class="fundo3">
	  <td width="100" height="29" align="center"><a href="consultas4.asp?Ordem=7"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;<font size="2"><b>FCG</b></font>&nbsp;<a href="consultas4.asp?Ordem=8"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
      <td width="170" height="29" align="center"><b><font size="2">Seção</td>
	  <td width="180" height="29" align="center"><b><font size="2">Esquadrão</td>      
      <td width="130" height="29" align="center"><b><font size="2">Marca</td>      
	  <td width="100" height="29" align="center"><b><font size="2">Situação</td>
	  <td width="60" height="29" align="center"><b><font size="2">Opções</td>	  
</tr>
<% If Session("Level") = 4 then
	Do While Not rsbanco3.EOF
		If rsbanco3("estabilizador_esquadrao") = Session("Esquadrao") then %>
<tr class="fundo5">
	  <td width="100" height="29" align="center"><font size="2"><% = rsbanco3("estabilizador_fcg") %></td>
      <td width="170" height="29" align="center"><font size="2"><% = rsbanco3("estabilizador_secao") %></td>
      <td width="180" height="29" align="center"><font size="2"><% = rsbanco3("estabilizador_esquadrao") %></td>      
      <td width="130" height="29" align="center"><font size="2"><% = rsbanco3("estabilizador_marca") %></td>
	  <td width="100" height="29" align="center"><font size="2"><% = rsbanco3("estabilizador_situacao") %></td>	
      <td width="60" height="29" align="center"><font size="2">
      <a href="alt_estabilizador.asp?codigo=<% = rsbanco3("estabilizador_codigo") %>"><img border="0" src="editar.gif" alt="Editar Estabilizador"></a>&nbsp;&nbsp;
      <% If Session("Level") < 3 Then %><a href="exc_estabilizador.asp?codigo=<% = rsbanco3("estabilizador_codigo") %>"><img border="0" src="del.gif" alt="Excluir Estabilizador"></a><% End If %></td>      
    </tr>
<%			rsbanco3.movenext
		Else
			rsbanco3.MoveNext
		End If
	Loop
Else
	Do While Not rsbanco3.EOF %>
<tr class="fundo5">
	  <td width="100" height="29" align="center"><font size="2"><% = rsbanco3("estabilizador_fcg") %></td>
      <td width="170" height="29" align="center"><font size="2"><% = rsbanco3("estabilizador_secao") %></td>
      <td width="180" height="29" align="center"><font size="2"><% = rsbanco3("estabilizador_esquadrao") %></td>      
      <td width="130" height="29" align="center"><font size="2"><% = rsbanco3("estabilizador_marca") %></td>
	  <td width="100" height="29" align="center"><font size="2"><% = rsbanco3("estabilizador_situacao") %></td>	
      <td width="60" height="29" align="center"><font size="2">
      <a href="alt_estabilizador.asp?codigo=<% = rsbanco3("estabilizador_codigo") %>"><img border="0" src="editar.gif" alt="Editar Estabilizador"></a>&nbsp;&nbsp;
      <% If Session("Level") < 3 Then %><a href="exc_estabilizador.asp?codigo=<% = rsbanco3("estabilizador_codigo") %>"><img border="0" src="del.gif" alt="Excluir Estabilizador"></a><% End If %></td>      
    </tr>
<%		rsbanco3.movenext
	Loop
End If %>	
</table>
<!-- Switch -->
<table width="740" align="center" border="1">
<tr>
	  <td class="fundo1" width="740" height="29" align="center" colspan="8"><b><font color="#7F0D11" size="2"><% = qtdswitch %></font><font size="2"> Switch(es)</td>	  
</tr>
<tr class="fundo3">
	  <td width="100" height="29" align="center"><a href="consultas4.asp?Ordem=9"><img border="0" src="seta.gif" alt="Crescente"></a>&nbsp;<font size="2"><b>FCG</b></font>&nbsp;<a href="consultas4.asp?Ordem=10"><img border="0" src="seta2.gif" alt="Decrescente"></a></td>
      <td width="150" height="29" align="center"><b><font size="2">Seção</td>
	  <td width="160" height="29" align="center"><b><font size="2">Esquadrão</td>      
      <td width="110" height="29" align="center"><b><font size="2">Marca</td>
	  <td width="80" height="29" align="center"><b><font size="2">Portas</td>         
	  <td width="80" height="29" align="center"><b><font size="2">Situação</td>
	  <td width="60" height="29" align="center"><b><font size="2">Opções</td>	  
</tr>
<% If Session("Level") = 4 then
	Do While Not rsbanco4.EOF
		If rsbanco4("switch_esquadrao") = Session("Esquadrao") then %>
<tr class="fundo5">
	  <td width="100" height="29" align="center"><font size="2"><% = rsbanco4("switch_fcg") %></td>
      <td width="150" height="29" align="center"><font size="2"><% = rsbanco4("switch_secao") %></td>
      <td width="160" height="29" align="center"><font size="2"><% = rsbanco4("switch_esquadrao") %></td>      
      <td width="110" height="29" align="center"><font size="2"><% = rsbanco4("switch_marca") %></td>
      <td width="80" height="29" align="center"><font size="2"><% = rsbanco4("switch_porta") %></td>      
	  <td width="80" height="29" align="center"><font size="2"><% = rsbanco4("switch_situacao") %></td>	
      <td width="60" height="29" align="center"><font size="2">
      <a href="alt_switch.asp?codigo=<% = rsbanco4("switch_codigo") %>"><img border="0" src="editar.gif" alt="Editar Switch"></a>&nbsp;&nbsp;
      <% If Session("Level") < 3 Then %><a href="exc_switch.asp?codigo=<% = rsbanco4("switch_codigo") %>"><img border="0" src="del.gif" alt="Excluir Switch"></a><% End If %></td>      
    </tr>
<%			rsbanco4.movenext
		Else
			rsbanco4.MoveNext
		End If
	Loop
Else
	Do While Not rsbanco4.EOF %>
<tr class="fundo5">
	  <td width="100" height="29" align="center"><font size="2"><% = rsbanco4("switch_fcg") %></td>
      <td width="150" height="29" align="center"><font size="2"><% = rsbanco4("switch_secao") %></td>
      <td width="160" height="29" align="center"><font size="2"><% = rsbanco4("switch_esquadrao") %></td>      
      <td width="110" height="29" align="center"><font size="2"><% = rsbanco4("switch_marca") %></td>
      <td width="80" height="29" align="center"><font size="2"><% = rsbanco4("switch_porta") %></td>      
	  <td width="80" height="29" align="center"><font size="2"><% = rsbanco4("switch_situacao") %></td>	
      <td width="60" height="29" align="center"><font size="2">
      <a href="alt_switch.asp?codigo=<% = rsbanco4("switch_codigo") %>"><img border="0" src="editar.gif" alt="Editar Switch"></a>&nbsp;&nbsp;
      <% If Session("Level") < 3 Then %><a href="exc_switch.asp?codigo=<% = rsbanco4("switch_codigo") %>"><img border="0" src="del.gif" alt="Excluir Switch"></a><% End If %></td>      
    </tr>
<%		rsbanco4.movenext
	Loop
End If %>	
</table>
<td align="center" width="340"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></td></form>
<!--#include file="rodape.asp"--></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>
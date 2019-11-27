<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then

	mes=CInt(request.querystring("numerario_mes"))

	extrato = ""
	If request.querystring("numerario_mes") = 1 then
		extrato = "Janeiro"
	Else
		If request.querystring("numerario_mes") = 2 then
			extrato = "Fevereiro"
		Else
			If request.querystring("numerario_mes") = 3 then
				extrato = "Março"
			Else
				If request.querystring("numerario_mes") = 4 then
					extrato = "Abril"
				Else
					If request.querystring("numerario_mes") = 5 then
						extrato = "Maio"
					Else
						If request.querystring("numerario_mes") = 6 then
							extrato = "Junho"
						Else
							If request.querystring("numerario_mes") = 7 then
								extrato = "Julho"
							Else
								If request.querystring("numerario_mes") = 8 then
									extrato = "Agosto"
								Else
									If request.querystring("numerario_mes") = 9 then
										extrato = "Setembro"
									Else
										If request.querystring("numerario_mes") = 10 then
											extrato = "Outubro"
										Else
											If request.querystring("numerario_mes") = 11 then
												extrato = "Novembro"
											Else
												extrato = "Desembro"
											End If
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from os order by os_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic
	
computador = 0
cartucho = 0
impressora = 0
login = 0
rede = 0
switch = 0
nobreak = 0
estabilizador = 0
monitor = 0
mouse = 0
teclado = 0

naoabertas = 0
abertas = 0
paralisadas = 0
fechadas = 0

gav = 0
binfa = 0
ec = 0
ei = 0
eie = 0
ep = 0
es = 0
gsb = 0
papv = 0
scoam = 0
dtceapv = 0

Do While Not rsbanco.EOF
	If rsbanco("os_mes") = mes And rsbanco("os_ano") = anoatual then
		If rsbanco("os_solicperiferico") = "Computador" then
			computador = computador + 1
		End If
		If rsbanco("os_solicperiferico") = "Cartucho" then
			cartucho = cartucho + 1
		End If
		If rsbanco("os_solicperiferico") = "Impressora" then
			impressora = impressora + 1
		End If
		If rsbanco("os_solicperiferico") = "Login/Senha" then
			login = login + 1
		End If
		If rsbanco("os_solicperiferico") = "Rede" then
			rede = rede + 1
		End If
		If  rsbanco("os_solicperiferico") = "Switch" then
			switch = switch + 1
		End If
		If rsbanco("os_solicperiferico") = "Nobreak" then
			nobreak = nobreak + 1
		End If
		If rsbanco("os_solicperiferico") = "Estabilizador" then
			estabilizador = estabilizador + 1
		End If
		If rsbanco("os_solicperiferico") = "Monitor" then
			monitor = monitor + 1
		End If
		If  rsbanco("os_solicperiferico") = "Mouse" then
			mouse = mouse + 1
		End If
		If rsbanco("os_solicperiferico") = "Teclado" then
			teclado = teclado + 1
		End If
		If rsbanco("os_status") = "0" then
			naoabertas = naoabertas + 1
		End If
		If rsbanco("os_status") = "1" then
			abertas = abertas + 1
		End If
		If rsbanco("os_status") = "2" then
			paralisadas = paralisadas + 1
		End If
		If rsbanco("os_status") = "3" then
			fechadas = fechadas + 1
		End If
		If rsbanco("os_solicesquadrao") = "2°/3° GAV" then
			gav = gav + 1
		End If
		If rsbanco("os_solicesquadrao") = "BINFA" then
			binfa = binfa + 1
		End If
		If rsbanco("os_solicesquadrao") = "EC" then
			ec = ec + 1
		End If
		If rsbanco("os_solicesquadrao") = "EI" then
			ei = ei + 1
		End If
		If rsbanco("os_solicesquadrao") = "EIE" then
			eie = eie + 1
		End If
		If  rsbanco("os_solicesquadrao") = "EP" then
			ep = ep + 1
		End If
		If rsbanco("os_solicesquadrao") = "ES" then
			es = es + 1
		End If
		If rsbanco("os_solicesquadrao") = "GSB" then
			gsb = gsb + 1
		End If
		If rsbanco("os_solicesquadrao") = "PAPV" then
			papv = papv + 1
		End If
		If  rsbanco("os_solicesquadrao") = "SCOAM" then
			scoam = scoam + 1
		End If
		If  rsbanco("os_solicesquadrao") = "DTCEA-PV" then
			dtceapv = dtceapv + 1
		End If			
		rsbanco.MoveNext
	Else
		rsbanco.MoveNext
	End If
Loop %>
<center>
<table border="1" width="600" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="510" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Controle de Ordem de Serviço</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Extrato Detalhado das Ordens de Serviço</b></font>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Mês de <font color="#FF0000"><% = extrato %></b></font></td>
    </tr>
</table>
  <table border="1" width="600" height="30">
    <tr class="fundo1">
      <td width="220" height="30" align="center" colspan="2"><b><font size="2">Números por Esquadrão</font></b></td>
      <td width="190" height="30" align="center" colspan="2"><b><font size="2">Tipos de serviços</font></b></td>
      <td width="190" height="30" align="center" colspan="2"><b><font size="2">Números e Ordem de Serviço</font></b></td>
    </tr>
    <tr>
      <td class="fundo3" width="190" height="30" align="right"><font size="2">Batalhão de Infantaria</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = binfa %></td>
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Computador</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = computador %></td>
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Não Abertas <img border="0" src="bolavermelha.gif"></font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = naoabertas %></td>
    </tr>
    <tr>
      <td class="fundo3" width="190" height="30" align="right"><font size="2">Esquadrão de Comando</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = ec %></td>    
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Cartucho</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = cartucho %></td>
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Abertas <img border="0" src="bolaverde.gif"></font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = abertas %></td>
    </tr>
    <tr>
      <td class="fundo3" width="190" height="30" align="right"><font size="2">Esquadrão de Intendência</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = ei %></td>    
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Impressora</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = impressora %></td>
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Andamento/Paralisadas <img border="0" src="bolaamarela.gif"></font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = paralisadas %></td>
    </tr>
    <tr>
      <td class="fundo3" width="190" height="30" align="right"><font size="2">Esquadrão de Infra-Estrutura</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = eie %></td>    
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Login/Senha</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = login %></td>
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Fechadas <img border="0" src="bolaazul.gif"></font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = fechadas %></td>
    </tr>
    <tr>
      <td class="fundo3" width="190" height="30" align="right"><font size="2">Esquadrão de Pessoal</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = ep %></td>    
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Rede</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = rede %></td>
      <td class="fundo1" width="190" height="30" align="center" colspan="2"><font size="2"><b>Unidades Apoiadas</td>
    </tr>
    <tr>
      <td class="fundo3" width="190" height="30" align="right"><font size="2">Esquadrão de Saúde</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = es %></td>    
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Switch</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = switch %></td>
      <td class="fundo3" width="160" height="30" align="center"><font size="2"><b>BAPV</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = (binfa + ec + ei + eie + ep + es + papv + scoam) %></td>   
    </tr>
    <tr>
      <td class="fundo3" width="190" height="30" align="right"><font size="2">Grupo de Serviço de Base</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = gsb %></td>    
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Nobreak</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = nobreak %></td>
      <td class="fundo3" width="160" height="30" align="center"><font size="2"><b>2º/3º GAV</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = gav %></td> 
    </tr>
    <tr>
      <td class="fundo3" width="190" height="30" align="right"><font size="2">Prefeitura de Aeronáutica de PV</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = papv %></td>    
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Estabilizador</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = estabilizador %></td>
      <td class="fundo3" width="160" height="30" align="center"><font size="2"><b>DTCEA-PV</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = dtceapv %></td> 
    </tr>
    <tr>
      <td class="fundo3" width="190" height="30" align="right"><font size="2">SCOAM</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = scoam %></td>    
      <td class="fundo3" width="160" height="30" align="right"><font size="2">Mouse/Teclado/Monitor</font></td>
      <td class="fundo5" width="30" height="30" align="center"><font color="#7F0D11" size="2"><b><% = (mouse + teclado + monitor) %></td>
      <td class="fundo1" width="160" height="30" align="center"><font size="2"><b>Total de Ordem de Serviço</td>
      <td width="30" height="30" align="center" bgcolor="#00FFFF"><font color="#7F0D11" size="2"><b><% = (naoabertas + abertas + paralisadas + fechadas) %></td>
    </tr>
  </table>
<form action="relatorio_os.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"><!--#include file="rodape.asp"--></form></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>
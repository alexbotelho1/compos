<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then

	secao=request.querystring("computador_secao")
	esquadrao=request.querystring("computador_esquadrao")
	tipo=request.querystring("computador_tipo")	
	so=request.querystring("computador_so")	
	situacao=request.querystring("computador_situacao")			

	set rsbanco=server.createobject("ADODB.Recordset")
		rsbanco.open "Select * from computador order by computador_esquadrao ASC",banco,AdOpenKeySet,AdLockOptimistic

formula = 0
If secao <> "1" then
	formula = formula + 1
End If
If esquadrao <> "2" then
	formula = formula + 2
End If
If tipo <> "4" then
	formula = formula + 4
End If
If so <> "8" then
	formula = formula + 8
End If
If situacao <> "16" then
	formula = formula + 16
End If


 %>
<body><center>
<table border="1" width="700" height="102">
    <tr>
      <td class="fundo1" width="100" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="600" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Consulta dos Cadastro de Hardware</b></font></td>
    </tr>
</table>
<% If (rsbanco.BOF And rsbanco.EOF) Or rsbanco.PageCount = 0 Then %>
<table border="0" cellpadding="0" cellspacing="0" width="740" align="center">		  
	<tr>
		<td align="center"><font face="Trebuchet MS" size="2" color="#ffffff"><i>N&atilde;o h&aacute; nada no momento!</i></font></td>
	</tr>
</table>
<form action="relatorio_os.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form>
<% Else %>	
<table width="700" align="center" border="1">
<tr>
	  <td class="fundo1" width="58" height="29" align="center"><b><font size="2">FCG</td>
      <td class="fundo1" width="112" height="29" align="center"><b><font size="2">Seção</td>
	  <td class="fundo1" width="130" height="29" align="center"><b><font size="2">Esquadrão</td>      
      <td class="fundo1" width="90" height="29" align="center"><b><font size="2">Tipo</td>
      <td class="fundo1" width="142" height="29" align="center"><b><font size="2">Sist. Operacional</td>
      <td class="fundo1" width="35" height="29" align="center"><b><font size="2">Qtd Proc</td>
      <td class="fundo1" width="57" height="29" align="center"><b><font size="2">Freq</td>
      <td class="fundo1" width="53" height="29" align="center"><b><font size="2">Mem</td>
	  <td class="fundo1" width="56" height="29" align="center"><b><font size="2">HD</font></td> 
</tr>
<% 	HowMany = 0
If formula = 0 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center" colspan="9">
	  	<table border="0" cellpadding="0" cellspacing="0" width="740" align="center">		  
			<tr>
				<td align="center"><font face="Trebuchet MS" size="2" color="#ffffff"><i>N&atilde;o h&aacute; nada no momento!</i></font></td>
			</tr>
			<tr>
				<td align="center"><form action="relatorio_os.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form></td>
			</tr>			
		</table>
	  </td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 1 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 2 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_esquadrao") = esquadrao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 3 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_secao") = secao And rsbanco("computador_esquadrao") = esquadrao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 4 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_tipo") = tipo then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><a href="computador.asp?codigo=<% = rsbanco("computador_codigo") %>"><font color="#7F0D11" size="2"><b><% = rsbanco("computador_fcg") %></a></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 5 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_tipo") = tipo And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 6 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_tipo") = tipo And rsbanco("computador_esquadrao") = esquadrao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 7 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_tipo") = tipo And rsbanco("computador_secao") = secao And rsbanco("computador_esquadrao") = esquadrao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 8 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_so") = so then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 9 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_so") = so And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 10 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_so") = so And rsbanco("computador_esquadrao") = esquadrao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 11 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_so") = so And rsbanco("computador_esquadrao") = esquadrao And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 12 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_so") = so And rsbanco("computador_tipo") = tipo then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 13 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_so") = so And rsbanco("computador_tipo") = tipo And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 14 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_so") = so And rsbanco("computador_tipo") = tipo And rsbanco("computador_esquadrao") = esquadrao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 15 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_so") = so And rsbanco("computador_tipo") = tipo And rsbanco("computador_esquadrao") = esquadrao And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 16 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 17 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 18 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_esquadrao") = esquadrao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 19 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_esquadrao") = esquadrao And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 20 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_tipo") = tipo then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 21 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_tipo") = tipo And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 22 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_tipo") = tipo And rsbanco("computador_esquadrao") = esquadrao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 23 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_tipo") = tipo And rsbanco("computador_esquadrao") = esquadrao And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 24 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_so") = so then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 25 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_so") = so And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 26 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_so") = so And rsbanco("computador_esquadrao") = esquadrao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 27 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_so") = so And rsbanco("computador_esquadrao") = esquadrao And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 28 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_tipo") = tipo And rsbanco("computador_so") = so then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 29 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_tipo") = tipo And rsbanco("computador_so") = so And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 30 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_tipo") = tipo And rsbanco("computador_so") = so And rsbanco("computador_esquadrao") = esquadrao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If
If formula = 31 then
	Do While Not rsbanco.EOF		 
		If rsbanco("computador_situacao") = situacao And rsbanco("computador_tipo") = tipo And rsbanco("computador_so") = so And rsbanco("computador_esquadrao") = esquadrao And rsbanco("computador_secao") = secao then %>
<tr>
	  <td class="fundo5" width="58" height="29" align="center"><font size="2"><% = rsbanco("computador_fcg") %></td>
      <td class="fundo5" width="112" height="29" align="center"><font size="2"><% = rsbanco("computador_secao") %></td>
      <td class="fundo5" width="130" height="29" align="center"><font size="2"><% = rsbanco("computador_esquadrao") %></td>      
      <td class="fundo5" width="90" height="29" align="center"><font size="2"><% = rsbanco("computador_tipo") %></td>
      <td class="fundo5" width="142" height="29" align="center"><font size="2"><% = rsbanco("computador_so") %></td>
      <td class="fundo5" width="35" height="29" align="center"><font size="2"><% = rsbanco("computador_qp") %></td>
      <td class="fundo5" width="57" height="29" align="center"><font size="2"><% = rsbanco("computador_procvelo") %><% = rsbanco("computador_procfreq") %></td>
      <td class="fundo5" width="53" height="29" align="center"><font size="2"><% = rsbanco("computador_memovelo") %><% = rsbanco("computador_memocapa") %></td>
	  <td class="fundo5" width="56" height="29" align="center"><font size="2"><% = rsbanco("computador_hdtama") %><% = rsbanco("computador_hdcapa") %></td>	
</tr>
<%			HowMany = HowMany + 1
			rsbanco.MoveNext
		Else
			rsbanco.MoveNext
		End If
	Loop
End If %>
</table>
<% End If %>
<% If HowMany = 0 Then %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">		  
	<tr>
		<td align="center"><font face="Trebuchet MS" size="2" color="#ffffff"><i>N&atilde;o h&aacute; nada no momento!</i></font></td>
	</tr>
</table>
<form action="javascript:history.go(-1)"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form>
<% End If %></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>
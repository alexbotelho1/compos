<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then %>
<body><center>
<table border="1" width="700" height="102">
    <tr>
      <td class="fundo1" width="90" height="102" align="center"><img border="0" src="logo.gif"></td>
      <td class="fundo2" width="610" height="102" align="center">
      <p style="margin-top: 0; margin-bottom: 0"><font color="#7F0D11" size="5"><b>Sistema de Inventário de Informática</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#070E5A"><b>Cadastrar Diversos</b></font></td>
    </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="700" height="35">
	<tr>
		<td width="400" height="35"><form method="GET" action="add11.asp">
			<table border="1" height="35">    
				<tr>
        			<td class="fundo1" width="110" height="35" align="center"><b>Periférico</b></td>
	        		<td class="fundo3" width="190" height="35" align="center"><input type="text" name="periferico_nome" size="19" style="border-style: inset; border-width: 5"></td>
   					<td class="fundo5" width="100" height="35" align="center"><input type="submit" value="Salvar"></td>
    			</tr>
			</table>
		</td></form>
		<td width="300" height="35"><form method="GET" action="exc8.asp">
			<table border="1" height="35">    
				<tr>
        			<td class="fundo3" width="200" height="35" align="center">
				<% 	Set rsbanco1 = Server.CreateObject("ADODB.Recordset")
					rsbanco1.Open "SELECT * FROM periferico ORDER BY periferico_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
						<select name="periferico_codigo">    
					<%  Do While Not rsbanco1.EOF %>								    
							<option value="<% = rsbanco1("periferico_codigo") %>"><% = rsbanco1("periferico_nome") %></option>
					<% 		rsbanco1.MoveNext 
						Loop %>											  
						</select>					  
				<% 	rsbanco1.Close
					Set rsbanco1= Nothing %>       
	        		</td>
					<td class="fundo5" width="100" height="35" align="center"><input type="submit" value="Excluir"></td>
    			</tr>
			</table>        
		</td></form>
	</tr>                
</table>
<table border="0" cellpadding="0" cellspacing="0" width="700" height="35">
	<tr>
		<td width="400" height="35"><form method="GET" action="add12.asp">
			<table border="1" height="35" width="400">    
				<tr>
        			<td class="fundo1" width="102" height="35" align="center"><b>Esquadrão</b></td>
	        		<td class="fundo3" width="183" height="35" align="center"><input type="text" name="esquadrao_nome" size="19" style="border-style: inset; border-width: 5"></td>
   					<td class="fundo5" width="93" height="35" align="center"><input type="submit" value="Salvar"></td>
    			</tr>
			</table>
		</td></form>
		<td width="300" height="35"><form method="GET" action="exc9.asp">
			<table border="1" height="35">    
				<tr>
        			<td class="fundo3" width="200" height="35" align="center">
				<% 	Set rsbanco2 = Server.CreateObject("ADODB.Recordset")
					rsbanco2.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
						<select name="esquadrao_codigo">    
					<%  Do While Not rsbanco2.EOF %>								    
							<option value="<% = rsbanco2("esquadrao_codigo") %>"><% = rsbanco2("esquadrao_nome") %></option>
					<% 		rsbanco2.MoveNext 
						Loop %>											  
						</select>					  
				<% 	rsbanco2.Close
					Set rsbanco2= Nothing %>       
	        		</td>
					<td class="fundo5" width="100" height="35" align="center"><input type="submit" value="Excluir"></td>
    			</tr>
			</table>        
		</td></form>
	</tr>                
</table>
<table border="0" cellpadding="0" cellspacing="0" width="700" height="35">
	<tr>
		<td width="400" height="35"><form method="GET" action="add18.asp">
			<table border="1" height="35" width="400">    
				<tr>
        			<td class="fundo1" width="102" height="35" align="center"><b>Seção</b></td>
	        		<td class="fundo3" width="182" height="35" align="center"><input type="text" name="secao_nome" size="19" style="border-style: inset; border-width: 5"></td>
   					<td class="fundo5" width="94" height="35" align="center"><input type="submit" value="Salvar"></td>
    			</tr>
			</table>
		</td></form>
		<td width="300" height="35"><form method="GET" action="exc15.asp">
			<table border="1" height="35">    
				<tr>
        			<td class="fundo3" width="200" height="35" align="center">
				<% 	Set rsbanco8 = Server.CreateObject("ADODB.Recordset")
					rsbanco8.Open "SELECT * FROM secao ORDER BY secao_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
						<select name="secao_codigo">    
					<%  Do While Not rsbanco8.EOF %>								    
							<option value="<% = rsbanco8("secao_codigo") %>"><% = rsbanco8("secao_nome") %></option>
					<% 		rsbanco8.MoveNext 
						Loop %>											  
						</select>					  
				<% 	rsbanco8.Close
					Set rsbanco8 = Nothing %>       
	        		</td>
					<td class="fundo5" width="100" height="35" align="center"><input type="submit" value="Excluir"></td>
    			</tr>
			</table>        
		</td></form>
	</tr>                
</table>
<table border="0" cellpadding="0" cellspacing="0" width="700" height="35">
	<tr>
		<td width="400" height="35"><form method="GET" action="add13.asp">
			<table border="1" height="35" width="400">    
				<tr>
        			<td class="fundo1" width="102" height="35" align="center"><b>Hardware</b></td>
	        		<td class="fundo3" width="183" height="35" align="center"><input type="text" name="hardware_nome" size="19" style="border-style: inset; border-width: 5"></td>
   					<td class="fundo5" width="93" height="35" align="center"><input type="submit" value="Salvar"></td>
    			</tr>
			</table>
		</td></form>
		<td width="300" height="35"><form method="GET" action="exc10.asp">
			<table border="1" height="35">    
				<tr>
        			<td class="fundo3" width="200" height="35" align="center">
				<% 	Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
					rsbanco3.Open "SELECT * FROM hardware ORDER BY hardware_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
						<select name="hardware_codigo">    
					<%  Do While Not rsbanco3.EOF %>								    
							<option value="<% = rsbanco3("hardware_codigo") %>"><% = rsbanco3("hardware_nome") %></option>
					<% 		rsbanco3.MoveNext 
						Loop %>											  
						</select>					  
				<% 	rsbanco3.Close
					Set rsbanco3= Nothing %>       
	        		</td>
					<td class="fundo5" width="100" height="35" align="center"><input type="submit" value="Excluir"></td>
    			</tr>
			</table>        
		</td></form>
	</tr>                
</table>
<table border="0" cellpadding="0" cellspacing="0" width="700" height="35">
	<tr>
		<td width="400" height="35"><form method="GET" action="add16.asp">
			<table border="1" height="35" width="400">    
				<tr>
        			<td class="fundo1" width="102" height="35" align="center"><b>Modelo Imp</b></td>
	        		<td class="fundo3" width="182" height="35" align="center"><input type="text" name="modimp_nome" size="19" style="border-style: inset; border-width: 5"></td>
   					<td class="fundo5" width="94" height="35" align="center"><input type="submit" value="Salvar"></td>
    			</tr>
			</table>
		</td></form>
		<td width="300" height="35"><form method="GET" action="exc13.asp">
			<table border="1" height="35">    
				<tr>
        			<td class="fundo3" width="200" height="35" align="center">
				<% 	Set rsbanco6 = Server.CreateObject("ADODB.Recordset")
					rsbanco6.Open "SELECT * FROM modimp ORDER BY modimp_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
						<select name="modimp_codigo">    
					<%  Do While Not rsbanco6.EOF %>								    
							<option value="<% = rsbanco6("modimp_codigo") %>"><% = rsbanco6("modimp_nome") %></option>
					<% 		rsbanco6.MoveNext 
						Loop %>											  
						</select>					  
				<% 	rsbanco6.Close
					Set rsbanco6 = Nothing %>       
	        		</td>
					<td class="fundo5" width="100" height="35" align="center"><input type="submit" value="Excluir"></td>
    			</tr>
			</table>        
		</td></form>
	</tr>                
</table>
<table border="0" cellpadding="0" cellspacing="0" width="700" height="35">
	<tr>
		<td width="400" height="35"><form method="GET" action="add14.asp">
			<table border="1" height="35" width="400">    
				<tr>
        			<td class="fundo1" width="102" height="35" align="center"><b>Marca Imp</b></td>
	        		<td class="fundo3" width="182" height="35" align="center"><input type="text" name="marcaimp_nome" size="19" style="border-style: inset; border-width: 5"></td>
   					<td class="fundo5" width="94" height="35" align="center"><input type="submit" value="Salvar"></td>
    			</tr>
			</table>
		</td></form>
		<td width="300" height="35"><form method="GET" action="exc11.asp">
			<table border="1" height="35">    
				<tr>
        			<td class="fundo3" width="200" height="35" align="center">
				<% 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
					rsbanco4.Open "SELECT * FROM marcaimp ORDER BY marcaimp_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
						<select name="marcaimp_codigo">    
					<%  Do While Not rsbanco4.EOF %>								    
							<option value="<% = rsbanco4("marcaimp_codigo") %>"><% = rsbanco4("marcaimp_nome") %></option>
					<% 		rsbanco4.MoveNext 
						Loop %>											  
						</select>					  
				<% 	rsbanco4.Close
					Set rsbanco4 = Nothing %>       
	        		</td>
					<td class="fundo5" width="100" height="35" align="center"><input type="submit" value="Excluir"></td>
    			</tr>
			</table>        
		</td></form>
	</tr>                
</table>
<table border="0" cellpadding="0" cellspacing="0" width="700" height="35">
	<tr>
		<td width="400" height="35"><form method="GET" action="add17.asp">
			<table border="1" height="35" width="400">    
				<tr>
        			<td class="fundo1" width="102" height="35" align="center"><b>Potência Nob</b></td>
	        		<td class="fundo3" width="182" height="35" align="center"><input type="text" name="nobreakpt_nome" size="19" style="border-style: inset; border-width: 5"></td>
   					<td class="fundo5" width="94" height="35" align="center"><input type="submit" value="Salvar"></td>
    			</tr>
			</table>
		</td></form>
		<td width="300" height="35"><form method="GET" action="exc14.asp">
			<table border="1" height="35">    
				<tr>
        			<td class="fundo3" width="200" height="35" align="center">
				<% 	Set rsbanco7 = Server.CreateObject("ADODB.Recordset")
					rsbanco7.Open "SELECT * FROM nobreakpt ORDER BY nobreakpt_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
						<select name="nobreakpt_codigo">    
					<%  Do While Not rsbanco7.EOF %>								    
							<option value="<% = rsbanco7("nobreakpt_codigo") %>"><% = rsbanco7("nobreakpt_nome") %></option>
					<% 		rsbanco7.MoveNext 
						Loop %>											  
						</select>					  
				<% 	rsbanco7.Close
					Set rsbanco7 = Nothing %>       
	        		</td>
					<td class="fundo5" width="100" height="35" align="center"><input type="submit" value="Excluir"></td>
    			</tr>
			</table>        
		</td></form>
	</tr>                
</table>
<table border="0" cellpadding="0" cellspacing="0" width="700" height="35">
	<tr>
		<td width="400" height="35"><form method="GET" action="add15.asp">
			<table border="1" height="35" width="400">    
				<tr>
        			<td class="fundo1" width="102" height="35" align="center"><b>Marca 
                    Nob</b></td>
	        		<td class="fundo3" width="182" height="35" align="center"><input type="text" name="marcanb_nome" size="19" style="border-style: inset; border-width: 5"></td>
   					<td class="fundo5" width="94" height="35" align="center"><input type="submit" value="Salvar"></td>
    			</tr>
			</table>
		</td></form>
		<td width="300" height="35"><form method="GET" action="exc12.asp">
			<table border="1" height="35">    
				<tr>
        			<td class="fundo3" width="200" height="35" align="center">
				<% 	Set rsbanco5 = Server.CreateObject("ADODB.Recordset")
					rsbanco5.Open "SELECT * FROM marcanb ORDER BY marcanb_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
						<select name="marcanb_codigo">    
					<%  Do While Not rsbanco5.EOF %>								    
							<option value="<% = rsbanco5("marcanb_codigo") %>"><% = rsbanco5("marcanb_nome") %></option>
					<% 		rsbanco5.MoveNext 
						Loop %>											  
						</select>					  
				<% 	rsbanco5.Close
					Set rsbanco5 = Nothing %>       
	        		</td>
					<td class="fundo5" width="100" height="35" align="center"><input type="submit" value="Excluir"></td>
    			</tr>
			</table>        
		</td></form>
	</tr>                
</table>
<table border="0" cellpadding="0" cellspacing="0" width="700" height="35">
	<tr>
		<td width="400" height="35"><form method="GET" action="add19.asp">
			<table border="1" height="35" width="400">    
				<tr>
        			<td class="fundo1" width="103" height="35" align="center"><b>STI</b></td>
	        		<td class="fundo3" width="180" height="35" align="center"><input type="text" name="sti_nomeguerra" size="19" style="border-style: inset; border-width: 5"></td>
   					<td class="fundo5" width="95" height="35" align="center"><input type="submit" value="Salvar"></td>
    			</tr>
			</table>
		</td></form>
		<td width="300" height="35"><form method="GET" action="exc16.asp">
			<table border="1" height="35">    
				<tr>
        			<td class="fundo3" width="200" height="35" align="center">
				<% 	Set rsbanco9 = Server.CreateObject("ADODB.Recordset")
					rsbanco9.Open "SELECT * FROM sti ORDER BY sti_antiguidade ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
						<select name="sti_codigo">    
					<%  Do While Not rsbanco9.EOF %>								    
							<option value="<% = rsbanco9("sti_codigo") %>"><% = rsbanco9("sti_nomeguerra") %></option>
					<% 		rsbanco9.MoveNext 
						Loop %>											  
						</select>					  
				<% 	rsbanco9.Close
					Set rsbanco9 = Nothing %>       
	        		</td>
					<td class="fundo5" width="100" height="35" align="center"><input type="submit" value="Excluir"></td>
    			</tr>
			</table>        
		</td></form>
	</tr>                
</table>
<form  action="admin.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"><!--#include file="rodape.asp"--></form></center></body></html>
<% Else
	Response.Redirect("admin.asp")
End If %>
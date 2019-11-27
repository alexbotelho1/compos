<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then %>
<center><body>
<table border="0" width="700" height="23" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" bordercolor="#000000">
	<tr>
    	<td width="97" height="101" align="center" style="border-left: 1px solid #000000; border-right-width: 1; border-top: 1px solid #000000; border-bottom-width: 1"><img border="0" src="logo.gif" width="73" height="90"></td>
    	<td width="603" height="101" align="center" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom-width: 1"><p style="margin-top: 0; margin-bottom: 0"><img border="0" src="aeronautica2.jpg" width="60" height="50"></p>
        <p style="margin-top: 0; margin-bottom: 0"><b><font size="4">COMANDO DA AERONÁUTICA</font></b></p>
        <p style="margin-top: 0; margin-bottom: 0">BASE AÉREA DE PORTO VELHO</p>
        <p style="margin-top: 0; margin-bottom: 0"><i><font size="2">Seção de Tecnologia da Informação</font></i></td>
    </tr>
</table>
<table border="0" width="695" height="23" bgcolor="#FFFFFF" bordercolor="#000000">
<tr>
<td class="fundo2" width="702" height="23">
 	<table border="1" width="680" height="244" bgcolor="#37832A" bordercolor="#000000">
		<tr>
			<form action="relatorio_mes.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo3" width="285" height="50" style="border-style: solid; border-width: 1; " align="center">
                  				<select size="1" name="os_mes">
  									<option value="1" selected>Janeiro</option>
  									<option value="2">Fevereiro</option>
  				  					<option value="3">Março</option>
  				  					<option value="4">Abril</option>
  				  					<option value="5">Maio</option>
  				  					<option value="6">Junho</option>
  				  					<option value="7">Julho</option>
  				  					<option value="8">Agosto</option>
  				  					<option value="9">Setembro</option>
  				  					<option value="10">Outubro</option>
  				  					<option value="11">Novembro</option>  				
  				  					<option value="12">Dezembro</option>				  								  				
  			  					</select>  
                  			</td> 
                		</tr>                		
                		<tr>                  
                  			<td class="fundo1" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Mês&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>
			<form action="relatorio_ano.asp">
				<td width="292" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo3" width="285" height="50" style="border-style: solid; border-width: 1; " align="center">
        	        			<select size="1" name="os_ano">
  				        			<option value="2007" selected>2007</option>
  				        			<option value="2008">2008</option>
  				        			<option value="2009">2009</option>
  				        			<option value="2010">2010</option>  						  								  				
  			        			</select>
                  			</td> 
                		</tr>                		
                		<tr>                  
                  			<td class="fundo1" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ano&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>      	          	                 	
        </tr>      		
		<tr>
			<form action="relatorio_periferico.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo3" width="285" height="50" style="border-style: solid; border-width: 1; " align="center">
	<% 	Set rsbanco5 = Server.CreateObject("ADODB.Recordset")
		rsbanco5.Open "SELECT * FROM periferico ORDER BY periferico_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_solicperiferico">								    
		<%  Do While Not rsbanco5.EOF %>								    
				<option value="<% = rsbanco5("periferico_nome") %>"><% = rsbanco5("periferico_nome") %></option>
		<% 		rsbanco5.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco5.Close
		Set rsbanco5= Nothing %>                  			
                  			</td> 
                		</tr>                		
                		<tr>                  
                  			<td class="fundo1" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Periférico&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>
			<form action="relatorio_esquadrao.asp">
				<td width="292" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo3" width="285" height="50" style="border-style: solid; border-width: 1; " align="center">
	<% 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
		rsbanco4.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="os_solicesquadrao">					    
		<%  Do While Not rsbanco4.EOF %>								    
				<option value="<% = rsbanco4("esquadrao_nome") %>"><% = rsbanco4("esquadrao_nome") %></option>
		<% 		rsbanco4.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco4.Close
		Set rsbanco4= Nothing %>
                  			</td> 
                		</tr>                  		
                		<tr>                  
                  			<td class="fundo1" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Esquadrão&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
            </form>          	           	                	
        </tr>
		<tr>
			<form action="numerario.asp">
				<td width="300" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo3" width="285" height="50" style="border-style: solid; border-width: 1; " align="center">
                  				<select size="1" name="numerario_mes">
  									<option value="1" selected>Janeiro</option>
  									<option value="2">Fevereiro</option>
  				  					<option value="3">Março</option>
  				  					<option value="4">Abril</option>
  				  					<option value="5">Maio</option>
  				  					<option value="6">Junho</option>
  				  					<option value="7">Julho</option>
  				  					<option value="8">Agosto</option>
  				  					<option value="9">Setembro</option>
  				  					<option value="10">Outubro</option>
  				  					<option value="11">Novembro</option>  				
  				  					<option value="12">Dezembro</option>				  								  				
  			  					</select>
                  			</td> 
                		</tr>                		
                		<tr>                  
                  			<td class="fundo1" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Extrato&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
          	</form>
			<form action="">
				<td width="292" height="61" align="center" style="border-style: none; border-width: medium">
              		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="285" height="50">
                		<tr>                  
                  			<td class="fundo3" width="285" height="50" style="border-style: solid; border-width: 1; " align="center">&nbsp;</td> 
                		</tr>                  		
                		<tr>                  
                  			<td class="fundo1" width="285" height="50" style="border-style: solid; border-width: 1; " align="center"><input type="submit" value="Futuras Ampliações" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td> 
                		</tr>
              		</table>
            	</td>
            </form>          	           	                	
        </tr>        
		<tr>
			<form action="relatorio_computador.asp">		
				<td width="709" height="60" align="center" style="border-style: none; border-width: medium" colspan="2">
					<table border="1" cellpadding="0" cellspacing="0" width="685" height="60">
						<tr>
      						<td class="fundo2" width="685" height="30" align="center" colspan="5"><b><font color="#000000">Relatórios de Hardware</font></b></td>
    					</tr>
    					<tr>
      						<td class="fundo2" width="137" height="30" align="center"><b><font color="#000000">Seção</font></b></td>
      						<td class="fundo2" width="138" height="30" align="center"><b><font color="#000000">Esquadrão</font></b></td>
      						<td class="fundo2" width="135" height="30" align="center"><b><font color="#000000">Tipo</font></b></td>
      						<td class="fundo2" width="144" height="30" align="center"><b><font color="#000000">Sistema Operacional</font></b></td>
      						<td class="fundo2" width="131" height="30" align="center"><b><font color="#000000">Situação</font></b></td>
    					</tr>
    					<tr>
      						<td class="fundo3" width="137" height="30" align="center">
	<% 	Set rsbanco3 = Server.CreateObject("ADODB.Recordset")
		rsbanco3.Open "SELECT * FROM secao ORDER BY secao_nome ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="computador_secao">			
				<option value="1" selected>Selecione</option>					    
		<%  Do While Not rsbanco3.EOF %>								    
				<option value="<% = rsbanco3("secao_nome") %>"><% = rsbanco3("secao_nome") %></option>
		<% 		rsbanco3.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco3.Close
		Set rsbanco3= Nothing %>      						
      						</td>
      						<td class="fundo3" width="138" height="30" align="center">
	<% 	Set rsbanco4 = Server.CreateObject("ADODB.Recordset")
		rsbanco4.Open "SELECT * FROM esquadrao ORDER BY esquadrao_codigo ASC",banco,AdOpenKeySet,AdLockOptimistic %>					  
			<select name="computador_esquadrao">
				<option value="2" selected>Selecione</option>					    
		<%  Do While Not rsbanco4.EOF %>								    
				<option value="<% = rsbanco4("esquadrao_nome") %>"><% = rsbanco4("esquadrao_nome") %></option>
		<% 		rsbanco4.MoveNext 
			Loop %>											  
			</select>					  
	<% 	rsbanco4.Close
		Set rsbanco4= Nothing %>                  				     						
      						</td>
      						<td class="fundo3" width="135" height="30" align="center">
              					<select size="1" name="computador_tipo">
  				  					<option value="4" selected>Selecione</option>            
  				  					<option>Cliente</option>
  				  					<option>Servidor</option> 
  				  					<option>Grande Porte</option>  				  				 				            
  			  					</select>       						
      						</td>
      						<td class="fundo3" width="144" height="30" align="center">
                  				<select size="1" name="computador_so">
  				      				<option value="8" selected>Selecione</option>
  				      				<option>Linux</option>  				           
  				      				<option>Windows 95</option>
  				      				<option>Windows 98</option> 
  				      				<option>Windows ME</option>
  				      				<option>Windows NT</option>  				 				  				 				            
  				      				<option>Windows XP</option>
  				      				<option>Windows 2000</option>
  				      				<option>Windows Vista</option>
  			      				</select>        						
      						</td>
      						<td class="fundo3" width="131" height="30" align="center">
                  				<select size="1" name="computador_situacao">
  				      				<option value="16" selected>Selecione</option>
  				      				<option>Uso</option>  				           
  				      				<option>Manutenção</option>
  				      				<option>Sucata</option> 
  			      				</select>       						
      						</td>
    					</tr>    					
    					<tr>
      						<td class="fundo1" width="685" height="30" align="center" colspan="5"><input type="submit" value="Buscar" style="font-family: Verdana; font-size:10 pt; font-weight:bold">&nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" value="Limpar" name="BTlimpar" style="font-family: Verdana; font-size:10 pt; font-weight:bold"></td>
    					</tr>    
					</table>				
				</td>
			</form>				       	           	                	
        </tr>              
	</table>
</td>
</tr>
</table>	
<form action="admin.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;">
<center><!--#include file="rodape.asp"--></center></form></body></center></html>
<% Else
	Response.Redirect("admin.asp")
End If %>
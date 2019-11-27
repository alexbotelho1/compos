<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head>
<!--#include file="config.asp"-->
<!--#include file="styles.asp"-->
<% If Session("LoggedIn") = True Then
	codsolic = request.querystring("codsolic") %>
<center><body>
<form name="News" action="pesquisa_os2.asp" method="GET">
<table border="0" width="700" height="23" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr>
    	<td width="97" height="101" align="center" style="border-left: 1px solid #000000; border-right-width: 1; border-top: 1px solid #000000; border-bottom-width: 1"><img border="0" src="logo.gif" width="73" height="90"></td>
    	<td width="603" height="101" align="center" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-top: 1px solid #000000; border-bottom-width: 1"><p style="margin-top: 0; margin-bottom: 0"><img border="0" src="aeronautica2.jpg" width="60" height="50"></p>
        <p style="margin-top: 0; margin-bottom: 0"><b><font size="4">COMANDO DA AERONÁUTICA</font></b></p>
        <p style="margin-top: 0; margin-bottom: 0">BASE AÉREA DE PORTO VELHO</p>
        <p style="margin-top: 0; margin-bottom: 0"><i><font size="2">Seção de Tecnologia da Informação</font></i></td>
    </tr>
</table>
<br><br>
<table border="1" width="347" height="27">
    <tr>
      <td class="fundo1" width="40" height="27">De:</td>
      <td class="fundo3" width="81" height="27">
        	<select size="1" name="os_dia">
  				<option value="1" selected>1</option>
  				<option value="2">2</option>
  				<option value="3">3</option>
  				<option value="4">4</option>
  				<option value="5">5</option>
  				<option value="6">6</option>
  				<option value="7">7</option>
  				<option value="8">8</option>
  				<option value="9">9</option>
  				<option value="10">10</option>
  				<option value="11">11</option>  				
  				<option value="12">12</option>
  				<option value="13">13</option>
  				<option value="14">14</option>
  				<option value="15">15</option>
  				<option value="16">16</option>
  				<option value="17">17</option>
  				<option value="18">18</option>
  				<option value="19">19</option>
  				<option value="20">20</option>
  				<option value="21">21</option>  				
  				<option value="22">22</option>
  				<option value="23">23</option>
  				<option value="24">24</option>
  				<option value="25">25</option>
  				<option value="26">26</option>
  				<option value="27">27</option>
  				<option value="28">28</option>
  				<option value="29">29</option> 
  				<option value="30">30</option>
  				<option value="31">31</option>  				  								  				
  			</select>      
      </td>
      <td class="fundo3" width="112" height="27">
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
      <td class="fundo3" width="104" height="27">
        	<select size="1" name="os_ano">
  				<option value="2007" selected>2007</option>
  				<option value="2008">2008</option>
  				<option value="2009">2009</option>
  				<option value="2010">2010</option>  						  								  				
  			</select>        
      </td>
    </tr>
</table>
<table border="1" width="347" height="27">
    <tr>
      <td class="fundo1" width="40" height="27">Até:</td>
      <td class="fundo3" width="81" height="27">
        	<select size="1" name="os_dia1">
  				<option value="1">1</option>
  				<option value="2">2</option>
  				<option value="3">3</option>
  				<option value="4">4</option>
  				<option value="5">5</option>
  				<option value="6">6</option>
  				<option value="7">7</option>
  				<option value="8">8</option>
  				<option value="9">9</option>
  				<option value="10">10</option>
  				<option value="11">11</option>  				
  				<option value="12">12</option>
  				<option value="13">13</option>
  				<option value="14">14</option>
  				<option value="15">15</option>
  				<option value="16">16</option>
  				<option value="17">17</option>
  				<option value="18">18</option>
  				<option value="19">19</option>
  				<option value="20">20</option>
  				<option value="21">21</option>  				
  				<option value="22">22</option>
  				<option value="23">23</option>
  				<option value="24">24</option>
  				<option value="25">25</option>
  				<option value="26">26</option>
  				<option value="27">27</option>
  				<option value="28">28</option>
  				<option value="29">29</option> 
  				<option value="30">30</option>
  				<option value="31" selected>31</option> 				  								  				
  			</select>      
      </td>
      <td class="fundo3" width="112" height="27">
        	<select size="1" name="os_mes1">
  				<option value="1">Janeiro</option>
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
  				<option value="12" selected>Dezembro</option>				  								  				
  			</select>      
      </td>
      <td class="fundo3" width="104" height="27">
        	<select size="1" name="os_ano1">
  				<option value="2007">2007</option>
  				<option value="2008">2008</option>
  				<option value="2009">2009</option>
  				<option value="2010" selected>2010</option>  						  								  				
  			</select>        
      </td>
    </tr>
</table>
<br>
<input type="submit" value="Pesquisar" name="BTincluir"></form>
<form action="admin.asp"><input type="submit" value="&nbsp;&nbsp;&nbsp;Voltar&nbsp;&nbsp;&nbsp;"></form>
</body></center></html>
<% Else
	Response.Redirect("admin.asp")
End If %>
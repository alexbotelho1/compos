<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head><body>
<!--#include file="config.asp"-->
<%
	registro = trim(request("exclui"))
		
	If registro = "" then
		response.write "Voc� n�o entrou com valor para exclus�o!"
	Else
		rsbanco.movefirst
			codigo = rsbanco("cod_sol")
		While ((codigo <> int(registro)) and (not rsbanco.eof))
			rsbanco.movenext
				If (not rsbanco.eof) then
					codigo = rsbanco("cod_sol")
				End If
		Wend
		If codigo = int(registro) then
			rsbanco.delete
			response.write "Registro exclu�do com Sucesso!!!"
		Else
			response.write "Registro exclu�do com sucesso!!!"
		End If
	End If
	
	set banco = nothing
	set rsbanco = nothing
	
	Response.Redirect("deletar.asp")
%>
</body></html>
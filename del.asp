<html><head><title>.:.:.:. COMPOS WEB .:.:.:.</title></head><body>
<!--#include file="config.asp"-->
<%
	registro = trim(request("exclui"))
		
	If registro = "" then
		response.write "Você não entrou com valor para exclusão!"
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
			response.write "Registro excluído com Sucesso!!!"
		Else
			response.write "Registro excluído com sucesso!!!"
		End If
	End If
	
	set banco = nothing
	set rsbanco = nothing
	
	Response.Redirect("deletar.asp")
%>
</body></html>
<%
	Dim includeValidacao : includeValidacao = 1
	Dim setDebugUser : setDebugUser = 0
	Sub setPathPage(pathPage)
		Session("setPathPage") = pathPage
	End Sub
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'	CLASSE ERRO
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	class ErrosDebug
		
		private dicErros
		private contadorUser
		private contadorDebug
			
		'Construtor
		Private Sub class_initialize()
			Set dicErros = Server.CreateObject("Scripting.Dictionary")
			contadorUser = 0
			contadorDebug = 0
		End Sub
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'	PRIVATE
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'Adicina erros ocorridos no dicionário
		Private Sub addErrorInterno(msg)
			If ( CStr(msg) = "" OR IsNull(msg)) Then
				Err.Raise 507, "ClasseErros", "[Classe Erros:addErrorInterno] Mensagem de erro vazia/nula."
			Else
				contadorDebug = contadorDebug + 1
				dicErros.add "debug" & contadorDebug, msg		
			End If		
		End Sub
	
		'Lista os erros ocorridos
		Private Sub listErrors()
			response.Clear()
			response.write("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
			response.write("<html xmlns=""http://www.w3.org/1999/xhtml"">")
			response.write("<head>")
			response.write("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />")
			response.write("<title>Erro</title>")
			response.write("<style>")
			response.write("html,body,center{height:100%;}")
			response.write("#conteudo{	height:auto !important;height:100%;min-height:100%;}")
			response.write("</style>")
			response.write("</head>")
			response.write("<body style=""background:#ccc;margin:0;"">")
			response.write("<center>")
			response.write("<div id=""conteudo"" style=""background:#fff;width:750px; text-align:left;padding:20px;"">")
			response.write("<div style=""background:url(https://www.site.com.br/homologacao/images/exclamationShield.png) left center no-repeat;padding-left:120px;height:140px;""><h1 style=""color:#000;font-size:14pt;"">Corporation</h1><p style=""color:#f00"">Atenção: Os campos abaixo estão com os valores inapropriados para a efetiva execução do sistema, por favor, realize a alteração e tente novamente</p></div>")
			response.write("<hr noshade=""noshade""/>")
			response.write("<div style=""font-size:10pt; font-family:verdana"">")
			
			'response.write("Path: " & session("setPathPage"))
			'response.write("Modulo: " & session("modulo_log"))
			For i = 1 to contadorUser
				response.write("<p>- Erro (" & i & "): " &  dicErros.Item("user"&i) & "</p>")
			Next
	
			if ( CLng(setDebugUser) > 0 AND CLng(setDebugUser) = CLng(Session("cd_user")) ) Then
				Response.Write("<br /><br /><span style=""color:red;"">Debug</span><br/>")
				For i = 1 to contadorDebug
					Response.write("<p>- Erro (" & i & "): " &  dicErros.Item("debug"&i))
					if Session("sql") <> "" Then 
						Response.Write("<br /><br />- SQL:<br />"& Session("sql"))
					End If
					Response.Write("</p>")
				Next		
			End If
			
			response.write("</div>")
			response.write("<hr noshade=""noshade""/>")
			response.write("<p><a id=""backLnk"" href=""javascript:void(0);"">Clique aqui para voltar e corrigir</a></p>")
			response.write("</div>")
			response.write("</center>")
			response.write("<script>")
			response.write("document.getElementById(""backLnk"").onclick = function(){window.history.go(-1);}")
			response.write("</script>")
			response.write("</body>")
			response.write("</html>")
			Response.End
		End Sub
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'	PUBLIC
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'Adicina erros ocorridos no dicionário
		Public Sub addError(msg)
			If ( CStr(msg) = "" OR IsNull(msg)) Then
				Err.Raise 507, "ClasseErros", "[Classe Erros:addErrorInterno] Mensagem de erro vazia/nula."
			Else
				contadorUser = contadorUser + 1
				dicErros.add "user" & contadorUser, msg	
			End If		
		End Sub
	
		'Função que executa a finalização dos objetos utilizados
		Public Sub closing
			dicErros = Nothing
		End Sub	
		
		'Verifica se ocorreu algum erro, caso tenha ocorredo lista os erros e para a execução.
		Public Sub isError()
			if ( contadorUser > 0 OR (contadorDebug > 0 AND setDebugUser <> 0) ) Then
				listErrors()
			End If 
		End Sub
		
		'Verifica se o usuario é debug
		Public Function verifyDebugUser
			Dim retorno : retorno = false
			if setDebugUser <> 0 AND setDebugUser = CLng(Session("cd_user")) Then
				retorno = true
			End If
			verifyDebugUser = retorno
		End Function
		
		'Mostra mensagem para usuário debug
		Public Sub viewDebugUser()
			if verifyDebugUser() AND Err.Number <> 0 Then
				Dim objErr
				Set objErr=Server.GetLastError()
				addErrorInterno("[ErrosDebug][viewDebugUser]{" & Session("setPathFile") & "}: <br/>Description: " & Err.Description & "<br/>Source: " & Err.Source)
				On Error GoTo 0
				isError() 
			End If
		End Sub
	
	End Class
	Set objErros = New ErrosDebug
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'	CLASSE CONSOLE
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	class Console
	
		private dicConsole
		private contadorConsole
		private dicParametros
		
		'Construtor
		Private Sub class_initialize()
			Set dicConsole = Server.CreateObject("Scripting.Dictionary")
			Set dicParametros = Server.CreateObject("Scripting.Dictionary")
			contadorConsole = 0
		End Sub
	
		'Quebra a string em pedaços para que possa ser melhor visualizada na tela de debug
		Private Function quebrarString(str)
			If ( isNull(str) ) Then 
				quebrarString = ""
			Else
				Dim RegEx
				Dim retorno : retorno = ""
				Dim contador : contador = 0
				Set RegEx = New regexp
				RegEx.Pattern = ".{0,80}"
				RegEx.Global = True
				RegEx.IgnoreCase = True
		
				Set Matches = RegEx.Execute(cStr(str))
				For Each Match in Matches
					If contador > 0 Then
						retorno = retorno & "<br />" &  Match
					Else
						retorno = retorno & Match
					End If
					contador = contador + 1
				Next	
				quebrarString = retorno
			End If
		End Function 
	
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'	PUBLIC
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
		'Finaliza os objetos utilizados pela classe
		Public Sub closing
			Set dicConsole = Nothing
			Set dicParametros = Nothing
		End Sub
	
		'Para a execução, limpa o buffer e mostra as mensgens de console para o usuário
		Public Sub console()
			If ( objErros.verifyDebugUser() ) Then
				Response.Clear
				Response.Write("<pre><hr>###Pagina###<br />" & Request.ServerVariables("PATH_INFO") & "<hr>")
				Response.Write("<hr>###Headers###<br />" & Request.ServerVariables("ALL_HTTP") & "<hr>")
				Response.Write("###Console### - Quantidade: " & dicConsole.Count & "<hr>")
				For i = 1 to dicConsole.Count
					Response.Write(i & " - " & dicConsole.Item(i))
				Next
				If ( dicParametros.Count > 0 ) Then
					Response.Write("<hr>###Parametros### - Quantidade: "& dicParametros.Count &"<hr>")
					Response.Write("<table style=""width:700px;border:1px solid black;border-collapse:collapse;"">")
					Response.Write("<tr style=""text-align:center;background-color:black;color:white;""><td style=""border: 1px solid black;"">Nome</td><td style=""border: 1px solid black;"">Tipo</td><td style=""border: 1px solid black;"">Nullable</td><td style=""border: 1px solid black;"">Valor</td><td style=""border: 1px solid black;"">Tamanho</td></tr>")
					Dim chaves : chaves = dicParametros.Keys
					For i = 0 to dicParametros.Count-1
						Dim color : color = "#FFF"
						IF (i MOD 2 = 0 ) Then color = "#CCC"
						Response.Write("<tr style=""boder:1px solid black;background-color:" & color & ";"">")
						Response.Write("<td style=""border: 1px solid black;"">" & chaves(i) & "</td>")
						Response.Write("<td style=""border: 1px solid black;text-align:center;"">" & dicParametros.Item(chaves(i))(1) & "</td>")
						Response.Write("<td style=""border: 1px solid black;text-align:center;"">" & dicParametros.Item(chaves(i))(2) & "</td>")
						Response.Write("<td style=""border: 1px solid black;"">" & quebrarString(dicParametros.Item(chaves(i))(3)) & "</td>")
						Response.Write("<td style=""border: 1px solid black;text-align:center;"">" & dicParametros.Item(chaves(i))(4) & "</td>")
						Response.Write("</tr>")
					Next
					Response.Write("</table>")
				End If
				Response.End
			End If
		End Sub
	
		'Grava uma mensagem que sera exibida quando o método console for chamado.
		Public Sub printh(str)
			If (objErros.verifyDebugUser) Then On Error Resume Next	
			If ( CStr(str) = "" OR IsNull(str)) Then
				Err.Raise 507, "ClasseErros", "[Classe Erros:addErrorInterno] Mensagem de erro vazia/nula."
			Else
				contadorConsole = contadorConsole + 1
				dicConsole.Add contadorConsole, str & "<br/>"
			End If
			objErros.viewDebugUser
		End Sub
	
		'Grava uma mensagem que sera exibida no navegador e quando o método console for chamado.
		Public Sub print(str)
			If (objErros.verifyDebugUser) Then 
				On Error Resume Next	
				If ( CStr(str) = "" OR IsNull(str)) Then
					Err.Raise 507, "ClasseErros", "[Classe Erros:addErrorInterno] Mensagem de erro vazia/nula."
				Else
					Response.Write(str & "<br />")
				End If
			End If
			If ( CStr(str) = "" OR IsNull(str)) Then
				Err.Raise 507, "ClasseErros", "[Classe Erros:addErrorInterno] Mensagem de erro vazia/nula."
			Else
				contadorConsole = contadorConsole + 1
				dicConsole.Add contadorConsole, str & "<br/>"
			End If
			objErros.viewDebugUser
		End Sub
	
		'Configura um objeto com os parâmetros no objeto de comando.
		Public Sub setParameters(parametros)
			If (objErros.verifyDebugUser) Then On Error Resume Next
			Set dicParametros = parametros
			objErros.viewDebugUser
		End Sub
	
	End Class
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'	CLASSE CONEXAO
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Class Conexao
		
		private connString
		private connClassValidacao
		private dicConnections
		
		'Construtor
		Private Sub class_initialize()
			SET dicConnections = Server.CreateObject("Scripting.Dictionary")
			fillConnections()
			connString = dicConnections.Item("NameDB1")
			Set connClassValidacao = Server.CreateObject("ADODB.Connection")
			connClassValidacao.ConnectionTimeOut = 60*5
		End Sub
		
		Private Sub fillConnections()
			dicConnections.add "NameDB1","Driver={SQL Server};Server=11.111.111.1;Database=database_name;UID=user;PWD=pass;"
			dicConnections.add "NameDB2","Driver={SQL Server};Server=11.111.111.1;Database=database_name;UID=user;PWD=pass;"
		End Sub
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'	PUBLIC
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'Função que executa a finalização dos objetos utilizados
		Public Sub closing
			Set connClassValidacao = Nothing
		End Sub
		
		' Pegar erros do objeto de conexão.
		Public Function countErrosConnection()
			getErrosConnection = connClassValidacao.Errors.Count
		End Function
		
		'Abre e retorna a conexao
		Public Function getConnection()	
			if Not IsNull(connClassValidacao) Then	
				If (objErros.verifyDebugUser) Then On Error Resume Next	
				Set connClassValidacao = Server.CreateObject("ADODB.Connection")
				connClassValidacao.Open connString
				Set getConnection = connClassValidacao
				objErros.viewDebugUser
			End If
		End Function
			
		'Retornar string configurada para conectar ao banco
		Public Function getConnectionString()
			getConnectionString = connString
		End Function
	
		'Inicia uma transaction
		Public Sub setBeginTrans()
			If ( objErros.verifyDebugUser) Then On Error Resume Next
			connClassValidacao.BeginTrans
			objErros.viewDebugUser
		End Sub
		
		'Seta o timeout do comando
		Public Sub setCommandTimeOut(segundos)
			If ( objErros.verifyDebugUser) Then On Error Resume Next
			connClassValidacao.CommandTimeout = segundos
			objErros.viewDebugUser	
		End Sub
		
		'Efetiva a transaction
		Public Sub setCommitTrans()
			If ( objErros.verifyDebugUser) Then On Error Resume Next
			connClassValidacao.CommitTrans
			objErros.viewDebugUser		
		End Sub
		
		'Configura uma nova string de conexão
		Public Sub setConnectionString(nameConnection)
			If ( objErros.verifyDebugUser) Then On Error Resume Next
			If (nameConnection = "" OR IsNull(nameConnection)) Then		
				Err.Raise 507,"ClasseConexao","[ClasseConexao:setConnectionString] Nome de conexão vazia/nula."
			End If
			If ( NOT dicConnections.Exists(nameConnection) ) Then
				Err.Raise 507,"ClasseConexao","[ClasseConexao:setConnectionString] Nome de conexão não existe."
			End IF
			objErros.viewDebugUser			
			connString = dicConnections.Item(nameConnection)
		End Sub
		
		'Seta o timeout da conexão
		Public Sub setConnectionTimeOut(segundos)
			If (segundos = "" OR IsNull(segundos)) Then
				If ( objErros.verifyDebugUser) Then On Error Resume Next
					Err.Raise 507,"ClasseConexao","[ClasseConexao:setConnectionTimeOut] Parametro segundos vazio/nulo."
				objErros.viewDebugUser			
			End If
			If ( objErros.verifyDebugUser) Then On Error Resume Next
			connClassValidacao.ConnectionTimeout = segundos
			objErros.viewDebugUser		
		End Sub
		
		'Cancela a transação
		Public Sub setRollBackTrans()
			If ( objErros.verifyDebugUser) Then On Error Resume Next
			connClassValidacao.RollBackTrans
			objErros.viewDebugUser			
		End Sub
		
	End Class
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'	CLASSE PARAMETROS
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Class  Parametros
	
		private method
		private p_tipos 
		private dicParams
		
		'Construtor
		Private sub class_initialize()
			Set dicParams = Server.CreateObject("Scripting.Dictionary")
			Set p_tipos = Server.CreateObject("Scripting.Dictionary") 
			fillTipos()
		End Sub
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'	PRIVATE
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'''''''''''''
		'Adiciona parametros no dicionario dicParams
		Private sub addParam(campo, tipo, nulo, valor, tamanho)
			dicParams.add LCase(campo), Array(LCase(campo), tipo, nulo, valor, tamanho)		
		End Sub
	
		'remove caracteres perigosos <, >, espaços, etc
		'return string limpa
		'Obs: os conjuntos de caracteres "%uFF1C" e "%uff1e" são respectivamente os sinais < e >
		Private Function clearString(tipo, v_string)
			'Remover qualquer tag html ex: <b>, <script>, etc
			if tipo = LCase("adHtml") then
				v_string = replace(v_string, "<", "&lt;")
				v_string = replace(v_string, ">", "&gt;")
				v_string = replace(v_string, "%uff1c", "&lt;", 1, -1, vbTextCompare)
				v_string = replace(v_string, "%uff1e", "&gt;", 1, -1, vbTextCompare)
			end if
			if not isnull(v_string) then
				v_string = trim(v_string)
			end if
	
			clearString = v_string
		end function
	
		'Adiciona os tipos de dados validos de acordo com os valores do ADODB
		'Os valores comentados não devem ser inseridos pois podem causar falhos no funcionamento devido não passarem por validação
		'Quando houver inclusão de novos parâmetros deve-se incluí-los na classe Console a chave/valor invertido.
		Private Sub fillTipos()
			p_tipos.add LCase("adSmallInt"), 2	
			p_tipos.add LCase("adInteger"), 3	
			p_tipos.add LCase("adSingle"), 4	
			p_tipos.add LCase("adDouble"), 5	
			p_tipos.add LCase("adBoolean"), 11	
			p_tipos.add LCase("adTinyInt"), 16	
			p_tipos.add LCase("adBigInt"), 20	
			p_tipos.add LCase("adChar"), 129	
			p_tipos.add LCase("adDBTimeStamp"), 135	
			p_tipos.add LCase("adVarChar"), 200	
			p_tipos.add LCase("adCusTime"), 135
			p_tipos.add LCase("adCusDate"), 135
			p_tipos.add LCase("adCusTimeStamp"), 135
			p_tipos.add LCase("adHtml"), 200
			p_tipos.add LCase("adXml"), 200
			p_tipos.add LCase("adLongVarChar"), 201	
			'p_tipos.add "adCurrency", 6	
			'p_tipos.add "adDate", 7
			'p_tipos.add "adBSTR", 8	
			'p_tipos.add "adIDispatch", 9	
			'p_tipos.add "adError", 10	
			'p_tipos.add "adVariant", 12	
			'p_tipos.add "adIUnknown", 13	
			'p_tipos.add "adDecimal", 14	
			'p_tipos.add "adUnsignedTinyInt", 17	
			'p_tipos.add "adUnsignedSmallInt", 18	
			'p_tipos.add "adUnsignedInt", 19	
			'p_tipos.add "adUnsignedBigInt", 21	
			'p_tipos.add "adFileTime", 64	
			'p_tipos.add "adGUID", 72	
			'p_tipos.add "adBinary", 128	
			'p_tipos.add "adWChar", 130	
			'p_tipos.add "adNumeric", 131	
			'p_tipos.add "adUserDefined", 132	
			'p_tipos.add "adDBDate", 133	
			'p_tipos.add "adDBTime", 134	
			'p_tipos.add "adChapter", 136	
			'p_tipos.add "adPropVariant", 138	
			'p_tipos.add "adVarNumeric", 139	
			'p_tipos.add "adVarWChar", 202	
			'p_tipos.add "adLongVarWChar", 203	
			'p_tipos.add "adVarBinary", 204	
			'p_tipos.add "adLongVarBinary", 205	
		End Sub	
		
		'Retorna valor de acordo com o metodos informado
		Private Function getValue(c)
			Dim retorno : retorno = null
			if (LCase(method) = "post") then 
				retorno = request.form(c)
			elseif (LCase(method) = "get") then
				retorno = request.querystring(c)
			elseif (LCase(method) = "cookie") then 
				retorno = Request.Cookies(c)
			elseif (LCase(method) = "request") then
				retorno = Request(c)
			elseif (LCase(method) = "session") then
				retorno = Session(c)
			else
				If ( objErros.verifyDebugUser) Then On Error Resume Next
				Err.Raise 507, "ClasseParametros", "[ClasseParametros:getValue] O metodo de adicionar parametros recebeu um tipo de 'metodo' invalido. {"& method &"}"
				objErros.viewDebugUser
			End If
			getValue = retorno	
		End Function
	
		'Validar date, time ou datetime conforme o tipo de dado passado
		'return true ou false
		'RexExp: as expressões regulares verificam se os tipos de dados passados possuem valores coerentes (dia, mes, ano)
		Private Function validarData(campo, valor, tipo)
			Dim retorno : retorno = true
			Dim dia : dia = 0
			Dim mes : mes = 0
			Dim ano : ano = 0
			Set objRegExp = new RegExp
			
			if tipo = LCase("adCusDate") then
				with objRegExp
					.Pattern = "^([0-2][0-9]|3[01])[\/|-](0[0-9]|1[0-2])[\/|-]([0-9]{4})$" 'BR
					.IgnoreCase = true
					.Global = True
				end with
				
				if objRegExp.test(valor) then
					dia = objRegExp.replace(valor, "$1")
					mes = objRegExp.replace(valor, "$2")
					ano = objRegExp.replace(valor, "$3")
					retorno = validateDate(dia, mes, ano)
				else
					retorno = false
				end if
				
			elseif tipo = LCase("adCusTime") then
				with objRegExp
					.Pattern = "^([0-1][0-9]|2[0-3])\:([0-5][0-9])(\:([0-5][0-9]))?$" 'BR
					.IgnoreCase = true
					.Global = True
				end with
				retorno = objRegExp.test(valor)
			elseif tipo = LCase("adDBTimeStamp") then
				with objRegExp
					.Pattern = "^([0-2][0-9]|3[01])[\/|-](0[0-9]|1[0-2])[\/|-]([0-9]{4}) (([0-1][0-9]|2[0-3])\:([0-5][0-9])(\:([0-5][0-9]))?)$" 'BR
					.IgnoreCase = true
					.Global = True
				end with
				
				if objRegExp.test(valor) then
					dia = objRegExp.replace(valor, "$1")
					mes = objRegExp.replace(valor, "$2")
					ano = objRegExp.replace(valor, "$3")
					retorno = validateDate(dia, mes, ano)
				else
					retorno = false
				end if
			elseif tipo = LCase("adCusTimeStamp") then
				with objRegExp
					.Pattern = "^([0-2][0-9]|3[01])[\/|-](0[0-9]|1[0-2])[\/|-]([0-9]{4})( ([0-1][0-9]|2[0-3])\:([0-5][0-9])(\:([0-5][0-9]))?)?$" 'BR				
					.IgnoreCase = true
					.Global = True
				end with
				
				if objRegExp.test(valor) then
					dia = objRegExp.replace(valor, "$1")
					mes = objRegExp.replace(valor, "$2")
					ano = objRegExp.replace(valor, "$3")
					retorno = validateDate(dia, mes, ano)
				else
					retorno = false
				end if
			end if
			if retorno = false then
				objErros.addError ("Parametro {'" & campo & "'} nao contem uma data valida.")
			end if
	
			validarData = retorno
		End Function
		
		'Valida se o numero é inteiro (integer)
		'return true ou false
		Private Function validarInteiro(valor, campo)
			Dim retorno : retorno = cBool(1)
			Dim regEx : Set regEx = New RegExp 
			valor = trim(valor)
			regEx.Pattern = "^[0-9]+$" 
			'regEx.Pattern = "^[0-9\,\.]+$" 
			retorno = RegEx.Test(valor)
			if (Len(valor) > 16) then retorno = cBool(0)
			
			if not cbool(retorno) then
				objErros.addError ("Parametro {'" & campo & "'} configurado como tipo INTEIRO nao possui um valor valido.")
			end if
			validarInteiro = cbool(retorno)
		End Function	
	
		'Valida se o valor é numero: inteiro ou real de acordo com seu tipo de dado
		'return true ou false
		Private Function validarNumero(valor, tipo, campo)
			Dim retorno : retorno = true
	
			Select Case tipo
				case 2 'adSmallInt", 2	
					retorno = validarInteiro(valor, campo)
				case 3 'adInteger", 3	
					retorno = validarInteiro(valor, campo)
				case 16 'adTinyInt", 16	
					retorno = validarInteiro(valor, campo)
				case 20 'adBigInt", 20	
					retorno = validarInteiro(valor, campo)
				case 4 'adSingle", 4	
					retorno = validarReal(valor, campo)
				case 5 'adDouble", 5	
					retorno = validarReal(valor, campo)
				case 14 'adDecimal", 14	
					retorno = validarReal(valor, campo)
			End Select
			
			validarNumero = retorno
		End Function
	
		'Valida se o numero e real (float)
		'return true ou false
		Private Function validarReal(valor, campo)
			Dim retorno : retorno = cBool(1)
			Dim regEx : Set regEx = New RegExp 
			valor = trim(valor)
			regEx.Pattern = "^[0-9\,\.]+$" 
			retorno = RegEx.Test(valor)
			if (Len(valor) > 16) then retorno = cBool(0)
			
			if not cbool(retorno) then
				objErros.addError ("Parametro {'" & campo & "'} configurado como tipo REAL nao possui um valor valido.")
			end if
			validarReal = cbool(retorno)
		End Function
		
		'Verifica se existem scripts em tudo menos no tipo de dado adHtml e adXml
		'return true ou false
		'Obs: os conjuntos de caracteres "%uFF1C" e "%uff1e" são respectivamente os sinais < e >
		Private Function validarScript(valor, tipo)
			Dim retorno : retorno = true
			
			if tipo <> LCase("adHtml") and tipo <> LCase("adXml") then
				Dim objRegExp: Set objRegExp = new RegExp
				with objRegExp
					.Pattern = "<[^>].*>|%uFF1C[^%uff1e].*%uff1e"
					.IgnoreCase = true
					.Global = True
				end with
				
				if objRegExp.test(valor) then
					retorno = false
					objErros.addError ("HTML nao aceita scripts." & valor &"---"& tipo)
				end if
			end if
			validarScript = retorno
		End Function
	
		'Verifica se existe scripts dentro do cdata no tipo de dado AdXml
		'return true ou false
		'Obs: os conjuntos de caracteres "%uFF1C" e "%uff1e" são respectivamente os sinais < e >
		'RegExp: realiza a busca de valores que começam com <![CDATA[
		'E contenham scripts e tags html e finaliza com ]]>
		'Ex: <![CDATA[%uFF1Cscript%uFF1e alert();%uFF1C/script%uFF1e]]>
		'Ex: <![CDATA[<script>teste<'/script>]]>
		Private Function validarScriptXml(valor, tipo)
			Dim retorno : retorno = true
			if tipo = LCase("adXml") then
				Dim objRegExp: Set objRegExp = new RegExp
				with objRegExp
					.Pattern = "(?=<!\[.+\[)(.*)(<|%uFF1C)(.*)(>|%uFF1e)(.*)(<|%uFF1C)(/.*)(>|%uFF1e)(.*)(\]\]>)"
					.IgnoreCase = true
					.Global = True
				end with
				
				if objRegExp.test(valor) then
					retorno = false
					objErros.addError("Xml nao aceita scripts e tag HTML." & valor &"---"& tipo)
				end if
			end if
			validarScriptXml = retorno
		End Function
			
		'Validacao do item
		'Primeiramente remove os valores perigosos
		'Caso tudo seja validado é executado a função addParam
		Private Sub validate(campo, tipo, nulo, valor, tamanho)
			Dim validacao : validacao = true
			valor = clearString(tipo, valor)
			 'Se aceitar nulo e o valor for nulo ou "" atribui o valor como null(objeto)
			if validateIsNull(valor) and nulo then
				valor = null
			'Se não aceitar nulo e o valor for nulo ou "" seta validação como false e adiciona erro.
			elseif validateIsNull(valor) and not nulo then
				validacao = false	
				objErros.addError("Campo {"&campo&"} nao permite valores nulos!")
			else
				'Verificar se o dado é um numero caso seja verifica se e inteiro e real e valida.
				Call validateTypeSize(campo, tipo, tamanho, valor)
				validacao = validarNumero(valor, p_tipos(tipo), campo)
				validacao = validarData(campo, valor, tipo)
				validacao = validarScript(valor, tipo)
				validacao = validarScriptXml(valor, tipo)
			end if
			'Se a validação for true adiciona o parametro
			if ( validacao ) Then 
				addParam campo, p_tipos(tipo), nulo, valor, tamanho
			End If
			objErros.isError()
		End Sub
		
		'Verificação de dia e mes se o ano é bisexto etc.
		'return true ou false
		Private Function validateDate(dia, mes, ano)
			Dim retorno : retorno = true
			Dim resto : resto = ano mod 4
			
			Select Case mes
				Case 1,3,5,7,8,10,12
					If ( CInt(dia) > 31 ) Then retorno = false
				Case 4,6,9,11
					if ( CInt(dia) > 30 ) Then retorno = false
				Case 2
					if ( resto = 0 AND CInt(dia) > 29) Then
						retorno = false
					Elseif ( resto <> 0 AND CInt(dia) > 28) Then
						retorno = false
					End If				
				Case Else
					retorno = true
			End Select
			validateDate = retorno
		End Function
	
		'Valida se o valor é nulo
		'return true ou false
		Private Function validateIsNull(v)
			If ( IsNUll(v) OR v = "" ) Then
				validateIsNull = true
			Else
				validateIsNuLL = false
			End If
		End Function
		
		Private Sub validateTypeSize(campo, tipo, tamanho, valor)
			
			Select case p_tipos(tipo)
				case 200,201
					If ( len(valor) > tamanho ) AND tamanho <> -1 Then
						Err.Raise 507, "ClassParametros", "[Classe Parametros:validateTypeSize] {" & LCase(campo) & "} contém um valor (tamanho) maior do que o configurado."
					End If
				
					if (tamanho <> -1  and tamanho < 1) then 
						Err.Raise 507, "ClassParametros", "[Classe Parametros:validateTypeSize] {" & LCase(campo) & "} deve ter seu tamanho maior ou igual a 1 (um);"
					end if
					
					if len(valor) > 8000 then
						objErros.addError "{" & LCase(campo) & "} deve ter menos de 8000 caracteres."
					end if
				case 2,3,4,5,11,16,20,135
					if (tamanho <> 0) then 
						Err.Raise 507, "ClassParametros", "[Classe Parametros:validateTypeSize] {" & LCase(campo) & "} deve ter seu tamanho igual a 0 (zero);"
					end if
				case 129
					if (tamanho <> 1) then 
						Err.Raise 507, "ClassParametros", "[Classe Parametros:validateTypeSize] {" & LCase(campo) & "} deve ter seu tamanho igual a 1 (um);"
					end if
			End Select
			objErros.isError
		End Sub
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'	PUBLIC
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
		'Adiciona Paramentro no dicionario proveninte de um metodo (post,get...)	
		'Realiza a validação, se os dados não contém valores indevidos é adicionado ao dicionário dicParams
		Public Function addRequest(campo, tipo, nulo, tamanho, metodo, nome_metodo)
			method = metodo
			Dim valor : valor = getValue(nome_metodo)
			addValue campo, tipo, nulo, tamanho, valor
		End Function
		
		'Adiciona um parametro diretamente com seu valor
		'Realiza a validação, se os dados não contém valores indevidos é adicionado ao dicionário dicParams
		Public Function addValue(campo, tipo, nulo, tamanho, valor)
			If ( objErros.verifyDebugUser ) Then On Error Resume Next
			If ( LCase(campo) = "" OR IsNull(LCase(campo)) ) Then Err.Raise 507, "ClasseParametros","[ClasseParametros:addValue] Parametro campo nao pode ser nulo."
			if ( tipo = "" OR IsNull(tipo) OR Not p_tipos.Exists(LCase(tipo)) ) Then Err.Raise 507, "ClasseParametros","[ClasseParametros:addValue] Parametro tipo nao e valido."
			if ( CBool(nulo) <> true AND CBool(nulo) <> false ) Then Err.Raise 507, "ClasseParametros","[ClasseParametros:addValue] Parametro nulo nao e booleano."
			if ( CInt(tamanho) < -1 ) Then Err.Raise 507, "ClasseParametros","[ClasseParametros:addValue] Parametro tamanho nao pode ser menor que -1."
			If ( dicParams.Exists(LCase(campo)) ) Then
				Err.Raise 507, "ClasseParametros","[ClasseParametros:addValue] Parametro ja existente {" & LCase(campo) & "}"
			else
	
				validate LCase(campo), LCase(tipo), nulo, valor, tamanho
			end if
			objErros.viewDebugUser
		End Function
	
		'Função que executa a finalização dos objetos utilizados
		Public Sub closing
			Set dicParams = Nothing
			Set p_tipos = Nothing
		End Sub
		
		'Retorna o valor de um item que já foi adicionado de acordo com o indice desejado
		'Indeces: 0 campo, 1 tipo, 2 nulo, 3 valor, 4 tamanho
		Public Function getItem(campo, indice)
			If ( objErros.verifyDebugUser ) Then On Error Resume Next
			If ( (CStr(LCase(campo)) = "" OR IsNull(LCase(campo))) OR (CStr(indice) = "" OR IsNULL(indice))) Then
				Err.Raise 507, "ClasseParametros","[ClasseParametros:getItem] Campo com valor vazio/nulo. {" & LCase(campo) & "}"
			Elseif ( CInt(indice) >= 5 OR CInt(indice) < 0 ) Then
				Err.Raise 507, "ClasseParametros", "[ClasseParametros:getItem] Indice passado nao e valido. {" & indice & "}" 
			End If	
			objErros.isError
			if ( dicParams.Exists(LCase(campo)) ) Then	
				getItem = dicParams.Item(LCase(campo))(indice)
			Else
				Err.Raise 507,"ClasseParametros", "[ClasseParametros:getItem] Parametro Inexistente {" & LCase(campo) & "}"
			End If		
			objErros.viewDebugUser
		End Function
		
		'Retorna o valor de um item que já foi adicionado
		Public Function getItemValue(campo)
			If ( objErros.verifyDebugUser ) Then On Error Resume Next
			If ( CStr(LCase(campo)) = "" OR IsNull(LCase(campo))) Then
				Err.Raise 507, "ClasseParametros", "[ClasseParametros:getItemValue] Campo com valor vazio/nulo. {" & LCase(campo) & "}"
			End If	
			objErros.isError
			if ( dicParams.Exists(LCase(campo)) ) Then
				getItemValue = dicParams.Item(LCase(campo))(3)	
			Else
				Err.Raise 507,"ClasseParametros", "[ClasseParametros:getItemValue] Parametro Inexistente {" & LCase(campo) & "}"
			End If
			objErros.viewDebugUser
		End Function
		
		'Retonar um dicionário com os parâmetros adicionados.
		Public Function getParameters()
			Set getParameters = dicParams
		End Function
	
		'Função que remove todos os itens do dicionário de parâmetros
		Public Sub removeAll()
			dicParams.RemoveAll()
		End Sub	
		
		'Função que remove um item do dicionário de parâmetros
		Public Sub removeItem(campo)
			If ( objErros.verifyDebugUser ) Then On Error Resume Next
			If ( CStr(LCase(campo)) = "" OR IsNull(LCase(campo))) Then
				Err.Raise 507, "ClasseParametros","[ClasseParametros:removeItem] Campo com valor vazio/nulo. {" & LCase(campo) & "}"
			End If	
			objErros.isError
			if ( dicParams.Exists(LCase(campo)) ) Then	
				dicParams.Remove(LCase(campo))
			Else
				Err.Raise 507, "ClasseParametros","[ClasseParametros:removeItem] Parametro Inexistente {" & LCase(campo) & "}"
			End If	
			objErros.viewDebugUser	
		End Sub	
		
		'Altera Paramentro no dicionario proveninte de um metodo (post,get...)	
		'Antes realiza a validação, se os dados não contém valores indevidos 
		'e o parâmetro existe, seu valor é alterado.
		Public function setRequest(campo, tipo, nulo, tamanho, meth, nome_method)
			method = meth	
			Dim valor : valor = getValue(nome_method)
			setValue campo, tipo, nulo, tamanho, valor
		end function
	
		'Altera parametro, antes realiza a validação, se os dados não contém valores indevidos 
		'e o parâmetro existe, seu valor é alterado.
		Public function setValue(campo, tipo, nulo, tamanho, valor)
			If ( objErros.verifyDebugUser ) Then On Error Resume Next
			If not ( dicParams.Exists(LCase(campo)) ) Then
				Err.Raise 507, "ClasseParametros","[ClasseParametros:setValue] Parametro nao existe {" & LCase(campo) & "}"
			else
				dicParams.Remove(LCase(campo))
				addValue LCase(campo), tipo, nulo, tamanho, valor
			end if
			objErros.viewDebugUser
		end function
	End Class
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'	CLASSE SQLCOMMAND
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Class SqlCommand
	
		public setCursorLocation
		private paramsSqlCmd
		private stringSql
		private objCmd
		private newRs
		
		'Construtor
		private sub class_initialize()
			setCursorLocation = 3
			Set dicDados = Server.CreateObject("Scripting.Dictionary")
			Set objCmd = Server.CreateObject("ADODB.Command")
			Set newRs = Server.CreateObject("ADODB.Recordset")
			Set paramsSqlCmd = Server.CreateObject("Scripting.Dictionary") 
			objCmd.CommandType = 1
			objCmd.CommandTimeout = 60*30
		End Sub
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'	PRIVATE
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
		'Adiciona os parametros no sqlcommand
		'Organiza na ordem em que é inserido na string sql ex: #nome#
		Private Sub fillParameters()
			Dim RegEx
			Set RegEx = New regexp
			RegEx.Pattern = "#[A-Za-z0-9_]+#"
			RegEx.Global = True
			RegEx.IgnoreCase = True
			If ( RegEx.Test(stringSql) ) Then
		
				If (  paramsSqlCmd.Count <= 0 ) Then
					Err.Raise 507, "ClasseSqlCommand", "[ClasseSqlCommand:fillParameters] Nenhum parametro adicionado ao objeto."
				End If
				If ( objCmd.Parameters.Count > 0 ) Then
					Dim quant : quant = objCmd.Parameters.Count -1
					While quant >= 0 
						objCmd.Parameters.Delete(quant)
						quant = quant - 1
					Wend
				End If
	
				Dim string_sql_resolvida : string_sql_resolvida = RegEx.Replace(stringSql,"?")
				objCmd.CommandText = string_sql_resolvida
				Set Matches  = RegEx.Execute(stringSql)
				For Each Match in Matches
					Dim pNome : pNome = Replace(Match,"#","")
					
					If Not paramsSqlCmd.Exists(LCase(pNome)) Then
						Err.Raise 507,"ClasseSqlCommand","[ClasseSqlCommand:fillParameters] Nao e possivel incluir parametros inexistentes. {" & LCase(pNome) & "}"
					End If

					On Error Resume Next
					'Response.write("<br />Parametro: " & LCase(pNome) & " Tipo: " & paramsSqlCmd(pNome)(1) & " Tamanho: " & paramsSqlCmd(pNome)(4) & " Valor: " & paramsSqlCmd(pNome)(3))
					objCmd.Parameters.append objCmd.CreateParameter(LCase(pNome), paramsSqlCmd(LCase(pNome))(1), 1, paramsSqlCmd(LCase(pNome))(4), paramsSqlCmd(LCase(pNome))(3))
					If Err.Number <> 0 then				
						Err.Raise 507, "ClasseSqlCommand", "[ClasseSqlCommand:fillParameters] erro ao inserir o parametro (" & LCase(pNome) & ")"
					End If
					On Error GOTO 0
				Next	
			Else
				objCmd.CommandText = stringSql
			End If		
		End Sub
		
		'Adiciona os parametros no sqlcommand
		'Organiza na ordem em que é inserido na string sql ex: #nome#
		Private Function fillParametersSql()
			Dim sqlFill : sqlFill = stringSql
			Dim RegEx
			Set RegEx = New regexp
			RegEx.Pattern = "#[A-Za-z0-9_]+#"
			RegEx.Global = True
			RegEx.IgnoreCase = True
			If ( RegEx.Test(stringSql) ) Then
			
				If (  paramsSqlCmd.Count <= 0 ) Then
					Err.Raise 507, "ClasseSqlCommand", "[ClasseSqlCommand:fillParametersSql] Nenhum parametro adicionado ao objeto."
				End If
				If ( objCmd.Parameters.Count > 0 ) Then
					Dim quant : quant = objCmd.Parameters.Count -1
					While quant >= 0 
						objCmd.Parameters.Delete(quant)
						quant = quant - 1
					Wend
				End If
	
				Set Matches = RegEx.Execute(stringSql)
				For Each Match in Matches
					Dim pNome : pNome = Replace(Match,"#","")
					If Not paramsSqlCmd.Exists(LCase(pNome)) Then
						Err.Raise 507, "ClasseSqlCommand", "[ClasseSqlCommand:fillParametersSql] Nao e possivel incluir parametros inexistentes. { " & LCase(pNome) & "} "
					End If
					
					if isNull(paramsSqlCmd(LCase(pNome))(3)) then
						sqlFill = Replace(sqlFill,Match,"null")
					else
						If ( paramsSqlCmd(LCase(pNome))(1) = 200 or paramsSqlCmd(LCase(pNome))(1) = 129 or paramsSqlCmd(LCase(pNome))(1) = 201 ) Then
							sqlFill = Replace(sqlFill,Match,"'" & paramsSqlCmd(LCase(pNome))(3) & "'")
						Else
							sqlFill = Replace(sqlFill,Match,paramsSqlCmd(LCase(pNome))(3))
						End If
					end if
				Next			
				fillParametersSql = sqlFill
			Else
				fillParametersSql = stringSql
			End If
		End Function	
	
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'	PUBLIC
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'Função que executa a finalização dos objetos utilizados
		Public Sub closing
			Set objCmd = Nothing
			Set dicDados = Nothing
			Set paramsSqlCmd = Nothing
			Set newRs = Nothing
		End Sub
	
		'Executa o comando e retorna a quantidade de registros afetados
		'Recomenda-se utilizar para execuções do tipo delete, updade e insert
		'Caso for utilizado select retornara -1
		Public Function cmdExecute()        
			Dim retorno
			fillParameters()
			objErros.isError
			If ( objErros.verifyDebugUser ) Then On Error Resume Next
			objCmd.Execute retorno
			cmdExecute = retorno
			objErros.viewDebugUser
		End Function
		
		'Executa um comando SELECT e retorna a primeira célula da primeira linha e coluna do RecordSet.
		Public Function cmdExecuteScalar()
			fillParameters()
			objErros.isError
			If ( objErros.verifyDebugUser ) Then On Error Resume Next
			Set newRs = Server.CreateObject("ADODB.Recordset")
			Set newRs = objCmd.Execute()
			if ( newRs.state <> 0 ) Then
				cmdExecuteScalar = newRs(0)
			Else
				cmdExecuteScalar = ""
			End If
			objErros.viewDebugUser
		End Function	
		
		'Executa o comando e retorna a o ID do dado inserido no banco
		'Exclusivo para execução de INSERT
		Public Function cmdInsertGetId()
			fillParameters()
			If ( objErros.verifyDebugUser ) Then On Error Resume Next
			if inStr(lcase(objCmd.CommandText), "insert") = 0 then
				Err.Raise 507,"ClasseSqlCommand","[ClasseSqlCommand:cmdInsertGetId] String nao contem um comando de insercao valido. {" & objCmd.CommandText & "}"
			end if
			objErros.isError
			objCmd.CommandText = objCmd.CommandText & ";SELECT SCOPE_IDENTITY() AS ID"
			Set newRs = Server.CreateObject("ADODB.Recordset")
			Set newRs = objCmd.Execute
			cmdInsertGetId = newRs("ID")
			objErros.viewDebugUser
		End Function
		
		'Retorna um objeto "dicionario" com os parâmetros incluidos no objeto de comando
		Public Function getParameters()
			Set getParameters = paramsSqlCmd
		End Function
		
		'Executa o comando e returna um recordset
		'Recomenda-se utilizar somente execuções do tipo select
		Public Function getRecordSet()	
			objErros.isError
			If ( objErros.verifyDebugUser ) Then On Error Resume Next
			fillParameters()
			Set newRs = Server.CreateObject("ADODB.Recordset")
			newRs.ActiveConnection = objCmd.ActiveConnection
			newRs.CursorType = 3
			newRs.LockType = 3
			Set newRs = objCmd.Execute()
			objErros.viewDebugUser
			Set getRecordSet = newRs
		End Function
		
		'Retonar string incluída no objeto de comando
		Public Function getStringSql()
			objErros.isError()
			getStringSql = stringSql
		End Function
		
		'Retorna string com os parâmetros preenchido com os valores
		Public Function getStringSqlFill()
			objErros.isError()
			getStringSqlFill = fillParametersSql()
		End Function
	
		'Configura qual conexao sera utilizada
		Public Sub setConnection( p_conn)
			objCmd.ActiveConnection = p_conn
			objCmd.ActiveConnection.CursorLocation = setCursorLocation
		End Sub
		
		'Configura um objeto com os parâmetros no objeto de comando.
		Public Sub setParameters(parametros)
			Set paramsSqlCmd = parametros
		End Sub
		
		'Configura String que será executada pelo objeto de comando
		Public Sub setStringSql(strSql)
			stringSql = strSql
		End Sub
		
	End Class 
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'	TOKEN
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Class Token
	
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'	PUBLIC
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'Método que gerar uma tokken em formato string
		Public Function generateToken()
			Set paransGT = New Parametros
			Set connect = New Conexao
			Set sc = new SqlCommand
	
			paransGT.addValue "sessionID","adInteger",false,0,Session.SessionID
			paransGT.addRequest "cd_user", "adInteger", false, 0, "session", "cd_user"
			Set connClassValidacao = connect.getConnection()
			sc.setConnection connClassValidacao
			sc.setParameters paransGT.getParameters()
			sc.setStringSql "proc_exec_gerar_token_auth @sessionId = #sessionID#, @cdUsuario = #cd_user#"
			generateToken = sc.cmdExecuteScalar()
			
			sc.Closing
			paransGT.Closing
			connect.Closing
		end function
		
		'Método que gera um tokken já incluído em um input type hidden.
		Public Function generateTokenHidden()
			Set paransGT = New Parametros
			Set connect = New Conexao
			Set sc = new SqlCommand
			paransGT.addValue "sessionID","adInteger",false,0,Session.SessionID
			paransGT.addRequest "cd_user", "adInteger", false, 0, "session", "cd_user"
			Set connClassValidacao = connect.getConnection()
			sc.setConnection connClassValidacao
			sc.setParameters paransGT.getParameters()
			sc.setStringSql "proc_exec_gerar_token_auth @sessionId = #sessionID#, @cdUsuario = #cd_user#"
			Dim tk : tk = sc.cmdExecuteScalar()
			generateTokenHidden = "<input type=""hidden"" name=""token"" value=""" & tk & """/>"
			
			sc.Closing
			paransGT.Closing
			connect.Closing
		end function
		
		'Método que verifica se o token é valido
		Public Sub verifyToken(token)
			Set paransVT = New Parametros
			Set connect = New Conexao
			Set sc = new SqlCommand
			
			paransVT.addValue "sessionID","adInteger",false,0,Session.SessionID
			paransVT.addValue "token","adVarChar",false,32,token
			Set connToken = connect.getConnection()
			sc.setConnection connToken
			sc.setParameters paransVT.getParameters()
			sc.setStringSql "proc_exec_validate_token_auth @token = #token#, @sessionId = #sessionID#"
			Dim retorno : retorno = sc.cmdExecuteScalar()
			sc.Closing
			paransVT.Closing
			connect.Closing
	
			If ( objErros.verifyDebugUser ) Then On Error Resume Next
			if ( CBool(retorno) = False) Then 
				Err.Raise 507, "Token", "Token Invalido"
			End If
			objErros.viewDebugUser
		End Sub
	
	End Class
	Set objToken = New Token
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'	TOOLKIT
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Class ToolKit
		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'	PUBLIC
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'Retorna um html com o valor de resultados encontrados
		'Ex: Resultados: 1 - 20 de 157 
		Public Function showAmountRecords(rs)	
			if rs.PageCount > 1 then
				inicio = (rs.AbsolutePage - 1) * 20 +1
				if rs.pageCount = 1 then
					fim = rs.recordCount
				elseif rs.pageCount = rs.AbsolutePage then
					fim = rs.recordCount
				else
					fim = inicio+19
				end if
			else
				inicio = 1
				fim = rs.recordCount
			end if
			showAmountRecords = "<font size = '1' face='Verdana'>"&lb_resultados&": "&inicio &" - "&fim & " de "&rs.recordCount & " " & lb_registros&"</font>"
		End Function
		
		'verifica se usuario esta logado
		Public Sub verifyLogon()
			Dim tipoLogon
			Dim dicSites 
			SET dicSites = Server.CreateObject("Scripting.Dictionary")
			dicSites.add "/site1","logado"
			dicSites.Add "/site2","logado"
			dicSites.Add "/site3","logado"
			dicSites.add "/site4","var_usuario"
			
			Set parans = New Parametros
			parans.addValue "urlexec","adVarchar",true,100,request.ServerVariables("SCRIPT_NAME")		
			
			Dim chSites : chSites = dicSites.Keys
			For i = 0 to dicSites.Count-1
				If ( inStr(parans.getItemValue("urlexec"), LCase(chSites(i))) = 0 ) Then
					tipoLogon = dicSites(chSites(i))
				End If
			Next
		
			parans.addRequest "logado","adBoolean",true,0,"session","logado"
	
			if not parans.getItemValue("logado") and inStr(parans.getItemValue("urlexec"), "/implement/logon.asp") = 0 then
				session.Abandon()
				response.Redirect(URLHTTP&"new_logiN.asp?t_http=1&MSG=Conexão Perdida")
			end if
			parans.closing()
		End Sub
		
		Public Function VerifyUrlRedirect(stringUrl)
			Set regExpUrl = New RegExp
			regExpUrl.IgnoreCase = True
			regExpUrl.Global = True
			
			regExpUrl.Pattern = "^http(s:|:)"
			
			if regExpUrl.test(stringUrl) then
				regExpUrl.Pattern = "^http(s:|:)//(www.|)site\.com\.br\/"
				if not regExpUrl.test(stringUrl) then
					stringUrl = "http://www.site.com.br/site/home.asp?url_modulo=pg_logout"
				end if
			end if
			VerifyUrlRedirect = stringUrl
		End Function
		
		Public Function replaceNull(stringValue, searchValue, replaceValue)
			retorno = null
			if not isNull(stringValue) then
				retorno = replace(stringValue, searchValue, replaceValue)
			end if
			replaceNull = retorno
		End Function
		
	End Class
	Set objToolKit = New ToolKit
	objToolKit.verifyLogon()
	
%>

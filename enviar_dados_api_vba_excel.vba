'Exemplo de Consulta de API no Excel com VBA;
'Utilize a biblioteca VBA Json para realizar a leitura do objeto Json;
'Habilite o RunTime nas Bibliotecas do Excel;
'Salve o arquivo XLSM para que as macros funcionem;

Sub EnvioDadosPostViaVBA()

'Declaracao das variaveis
Dim cadeia As String
Dim Result() As String
Dim contador As Integer
Dim total, linha As Integer
Dim ws As Worksheet

'Selecionamos nossa planilha resultados
Set ws = Worksheets("matriculas")

'atribuimos a url consulta
Url = "SUA URL" 'nao esqueca de preencher com http ou https;

'Definimos os contadores
total = 0
linha = 2

'laco de repeticao funciona ate encontrar uma linha vazia;
Do While ws.Cells(linha, 1).Value <> ""

'atribui informacoes da planilha nas variaveis
idAluno = ws.Cells(linha, 1).Value
idEmpresa = ws.Cells(linha, 2).Value
idPerfil = ws.Cells(linha, 3).Value
idTreinamento = ws.Cells(linha, 4).Value
dataMatricula = ws.Cells(linha, 5).Value
dataValidade = ws.Cells(linha, 6).Value

Result = Split(idTreinamento, ",")

    For i = LBound(Result()) To UBound(Result())
    
        'Preenche e Envia requisicao POST com os dados em formato Json: atencao para a concatenacao da string com os elementos; #ponto critico de funcionamento
        JsonBody = "{""dominio"": ""seuDominio"",""senha"": ""suaSenha"",""classe"": ""matricula"",""metodo"": ""cadastrar"",""id_aluno"": """ & idAluno & """,""id_empresa"": """ & idEmpresa & """,""id_perfil"": """ & idPerfil & """,""id_treinamento"": """ & Result(i) & """,""data"": """ & dataMatricula & """,""hora"": """",""liberar"": ""1"",""origem"": ""0"",""validade"": """ & dataValidade & """,""solicitacao_rematricula"": ""0""}"
        
        'Criamos nosso objeto de requisicao
        Set objpostHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
        
        objpostHTTP.Open "POST", Url, False
        objpostHTTP.setRequestHeader "cache-control", "no-cache"
        objpostHTTP.setRequestHeader "Accept", "application/json"
        objpostHTTP.setRequestHeader "Content-Type", "application/json"

        'Enviamos nosso JsonBody para a API;
        objpostHTTP.Send (JsonBody)
        
        strResult = objpostHTTP.responseText
        json = strResult
        
        'convertemos a resposta da API com nossa biblioteca VBA Json;
        Set objetoJson = JsonConverter.ParseJson(json)
        
        'Pega o retorno do POST
        retorno = objetoJson("status")
        
        'Preenche o retorno na planilha
        ws.Cells(linha, 7 + i) = retorno
        
        'Aguarda o tempo da API: 5 requisicao a cada 20 segundos
        Application.Wait (Now + TimeValue("0:00:05"))
        
    Next i
    
    linha = linha + 1
    total = total + 1
    
Loop

MsgBox "#Automacao eh Bonito!" & vbNewLine & "#ProgramadorPreguicoso" & vbNewLine & vbNewLine & "Foram Matriculados " & total & " Alunos =)" & vbNewLine & vbNewLine & "Tarefa feita com excelencia!"

End Sub
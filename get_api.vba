Sub ListarAlunos()
Dim JsonBody As String
Dim objpostHTTP As Object
Dim json As String
Dim jsonObject As Object, item As Object
Dim i As Long
Dim ws As Worksheet
Dim resposta As Integer
Dim total As Integer

'atribuimos a url consulta
Url = "URL_API"

'Atribui a requisição com a autenticação
JsonBody = "{""dominio"": ""SeuDominioExemplo"",""senha"": ""ChaveAPI"",""classe"": ""aluno"",""metodo"": ""listar""}"
 
'Criamos nosso objeto de requisição e enviamos
Set objpostHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

objpostHTTP.Open "POST", Url, False
objpostHTTP.setRequestHeader "cache-control", "no-cache"
objpostHTTP.setRequestHeader "Accept", "application/json"
objpostHTTP.setRequestHeader "Content-Type", "application/json"
objpostHTTP.Send (JsonBody)

If objpostHTTP.Status <> 200 Then
    MsgBox "HTTP Error " & objpostHTTP.Status & ". Corrija o erro e tente novamente!"
    Exit Sub
Else
    resposta = MsgBox("Carregar TODOS os alunos?", vbQuestion + vbYesNo, "Servidor Status: " & objpostHTTP.Status & " = Connected!")
    
    If resposta = vbYes Then
        'MsgBox "Yes"
        'limpa planilha com os dados antigos
        Sheets("alunos").Cells.Clear
    Else
        'MsgBox "Noooo"
        Exit Sub
    End If
End If

strResult = objpostHTTP.responseText
json = strResult
 
Set objetoJson = JsonConverter.ParseJson(json)
 
'Selecionamos nossa planilha resultados
Set ws = Worksheets("alunos")
 
'Criamos as células de cabeçalho
ws.Cells(1, 1) = "EMAIL"
ws.Cells(1, 2) = "ID"
ws.Cells(1, 3) = "NOME"
ws.Cells(1, 4) = "SOBRENOME"
ws.Cells(1, 5) = "EMPRESA"
ws.Cells(1, 6) = "PERFIL"

'Identifica o cabeçalho da tabela
ws.Cells(1, 1).Font.Bold = True
ws.Cells(1, 2).Font.Bold = True
ws.Cells(1, 3).Font.Bold = True
ws.Cells(1, 4).Font.Bold = True
ws.Cells(1, 5).Font.Bold = True
ws.Cells(1, 6).Font.Bold = True
 
'Fazemos um loop na propriedade results da resposta da API
i = 2 'Começaremos o contador na linha 2
For Each t In objetoJson
    ws.Cells(i, 1) = t("login")
    ws.Cells(i, 2) = t("id")
    ws.Cells(i, 3) = t("nome")
    ws.Cells(i, 4) = t("sobrenome")
    ws.Cells(i, 5) = t("empresa")
    ws.Cells(i, 6) = t("perfil")
    i = i + 1
    total = i
Next

MsgBox "#Automação é Bonito!" & vbNewLine & "#ProgramadorPreguiçoso" & vbNewLine & vbNewLine & "Foram encontrados " & total - 2 & " Alunos =)" & vbNewLine & vbNewLine & "Tarefa feita com excelência!"
 
End Sub
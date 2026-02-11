' --- MACRO VBA VERSÃO 16 - COM ENVIO PARA API ONLINE ---
' Esta versão envia os dados via HTTP POST para uma API na nuvem
' CONFIGURAÇÃO: Altere a URL_API abaixo para o endereço do seu servidor hospedado

Sub ExportarDadosParaAPI()
    ' ========== CONFIGURAÇÃO ==========
    ' IMPORTANTE: Altere esta URL após fazer o deploy do servidor
    Const URL_API As String = "https://seu-servidor.com/api/dados"
    ' ==================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, count As Integer, j As Integer, k As Integer, c As Integer
    Dim json As String
    Dim sheetName As String
    Dim firstSheet As Boolean, firstRow As Boolean, firstDS As Boolean
    Dim dataHoje As Date
    Dim http As Object
    Dim response As String

    dataHoje = Date

    On Error GoTo ErrorHandler

    ' ========== CONSTRUÇÃO DO JSON (igual antes) ==========
    json = "{" & vbCrLf
    json = json & "  ""timestamp"": """ & Format(Now, "yyyy-mm-ddThh:mm:ss") & """," & vbCrLf
    json = json & "  ""analises"": [" & vbCrLf

    firstSheet = True

    ' 1. PROCESSAR CAPA (STATUS)
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Capa")
    On Error GoTo ErrorHandler

    If Not ws Is Nothing Then
        json = json & "    {" & vbCrLf
        json = json & "      ""nome"": ""Status Geral""," & vbCrLf
        json = json & "      ""tipo"": ""capa""," & vbCrLf
        json = json & "      ""status_data"": [" & vbCrLf

        Dim tanques As Variant, linhas As Variant
        tanques = Array("Desengraxante (T1)", "Ativação (T5)", "Ativação (T6)", "Níquel WATT (T10)", "Estanhagem A (T12)", "Estanhagem B (T13)")
        linhas = Array(6, 9, 12, 15, 18, 23)

        For i = 0 To UBound(tanques)
            Dim l As Long: l = linhas(i)
            Dim dtProximaVal As Variant
            Dim dtProximaStr As String, txtStatus As String

            dtProximaVal = ws.Cells(l, 10).Value
            If dtProximaVal = "" Then dtProximaVal = ws.Cells(l + 1, 10).Value

            ' Força o formato DD/MM/YYYY
            If IsDate(dtProximaVal) Then
                dtProximaStr = Format(dtProximaVal, "dd/mm/yyyy")
            Else
                dtProximaStr = ws.Cells(l, 10).Text
            End If

            If Not IsDate(dtProximaVal) Then
                txtStatus = "SEM DATA"
            ElseIf CDate(dtProximaVal) < dataHoje Then
                txtStatus = "TESTE ATRASADO"
            Else
                txtStatus = "OK"
            End If

            json = json & "        {""tanque"": """ & tanques(i) & """, "
            json = json & """proxima"": """ & dtProximaStr & """, "
            json = json & """status"": """ & txtStatus & """}"
            If i < UBound(tanques) Then json = json & ","
            json = json & vbCrLf
        Next i
        json = json & "      ]" & vbCrLf
        json = json & "    }"
        firstSheet = False
    End If

    ' 2. PROCESSAR GRÁFICOS ANALÍTICOS
    For Each ws In ThisWorkbook.Worksheets
        If (Left(ws.Name, 4) = "Ana." And InStr(ws.Name, "Adi.") = 0 And InStr(ws.Name, "SELANTE") = 0) Then
            If Not firstSheet Then json = json & "," & vbCrLf
            firstSheet = False

            sheetName = Replace(ws.Name, "Ana. ", "")
            sheetName = Replace(sheetName, "Ana.", "")
            If Trim(sheetName) = "Des" Then sheetName = "Desengraxante"

            json = json & "    {" & vbCrLf
            json = json & "      ""nome"": """ & Trim(sheetName) & """," & vbCrLf
            json = json & "      ""tipo"": ""analise""," & vbCrLf

            Dim colIndices As Collection: Set colIndices = New Collection
            Dim colLabels As Collection: Set colLabels = New Collection
            If InStr(ws.Name, "Tanque 5") > 0 Then
                colIndices.Add 5: colLabels.Add "pH lido"
            ElseIf InStr(ws.Name, "Níquel") > 0 Then
                colIndices.Add 8: colLabels.Add "Niº (g/L)": colIndices.Add 9: colLabels.Add "NiCl2 (mL/L)": colIndices.Add 10: colLabels.Add "H3BO3 (g/L)"
            ElseIf InStr(ws.Name, "Estanhagem") > 0 Then
                colIndices.Add 7: colLabels.Add "Sn (g/L)": colIndices.Add 8: colLabels.Add "H2SO4 (mL/L)"
            Else
                colIndices.Add 6: colLabels.Add "Concentração"
            End If

            json = json & "      ""ultimas_analises"": [" & vbCrLf
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
            Dim rowsToExport As Collection: Set rowsToExport = New Collection
            count = 0
            For i = lastRow To 5 Step -1
                If IsDate(ws.Cells(i, 1).Value) And ws.Cells(i, 1).Value <> "" Then
                    rowsToExport.Add i: count = count + 1
                End If
                If count = 4 Then Exit For
            Next i

            firstRow = True
            For i = rowsToExport.count To 1 Step -1
                If Not firstRow Then json = json & "," & vbCrLf
                firstRow = False
                Dim r As Long: r = rowsToExport(i)
                ' Força o formato DD/MM/YYYY na tabela
                json = json & "        {""Data"": """ & Format(ws.Cells(r, 1).Value, "dd/mm/yyyy") & """, ""Resp"": """ & ws.Cells(r, 2).Value & """"
                For k = 1 To colIndices.count
                    json = json & ", ""Val" & k & """: " & Replace(CStr(Val(ws.Cells(r, colIndices(k)).Value)), ",", ".")
                Next k
                json = json & "}"
            Next i
            json = json & vbCrLf & "      ]," & vbCrLf

            json = json & "      ""chart_data"": {" & vbCrLf
            json = json & "        ""labels"": ["
            For i = rowsToExport.count To 1 Step -1
                ' No gráfico mantemos DD/MM para não poluir o eixo
                json = json & """" & Format(ws.Cells(rowsToExport(i), 1).Value, "dd/mm") & """"
                If i > 1 Then json = json & ","
            Next i
            json = json & "]," & vbCrLf
            json = json & "        ""datasets"": ["
            firstDS = True
            For k = 1 To colIndices.count
                If Not firstDS Then json = json & ","
                firstDS = False
                json = json & "{""label"": """ & colLabels(k) & """, ""data"": ["
                For i = rowsToExport.count To 1 Step -1
                    json = json & Replace(CStr(Val(ws.Cells(rowsToExport(i), colIndices(k)).Value)), ",", ".")
                    If i > 1 Then json = json & ","
                Next i
                json = json & "]}"
            Next k
            json = json & "]" & vbCrLf
            json = json & "      }" & vbCrLf
            json = json & "    }"
        End If
    Next ws

    json = json & vbCrLf & "  ]" & vbCrLf & "}"

    ' ========== ENVIAR PARA API VIA HTTP POST ==========
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    ' Configura a requisição
    http.Open "POST", URL_API, False
    http.setRequestHeader "Content-Type", "application/json; charset=UTF-8"

    ' Envia os dados
    http.send json

    ' Verifica resposta
    If http.Status = 200 Then
        MsgBox "✅ Dados enviados com sucesso para a nuvem!" & vbCrLf & vbCrLf & _
               "Resposta: " & http.responseText, vbInformation, "WEG - Exportação"
    Else
        MsgBox "❌ Erro ao enviar dados!" & vbCrLf & vbCrLf & _
               "Status: " & http.Status & vbCrLf & _
               "Resposta: " & http.responseText, vbCritical, "WEG - Erro"
    End If

    Set http = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "❌ Erro durante a exportação:" & vbCrLf & vbCrLf & _
           "Erro: " & Err.Description & vbCrLf & _
           "Número: " & Err.Number, vbCritical, "WEG - Erro"
End Sub

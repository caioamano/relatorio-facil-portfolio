Attribute VB_Name = "ModuloRelatorioFacil"
Option Explicit

' =========================================================
' 1. CADASTRAR PRODUTO
' =========================================================

Public Sub CadastrarProduto()

    Dim wsForm As Worksheet
    Dim wsBase As Worksheet
    Dim tbl As ListObject
    Dim novaLinha As ListRow

    Dim sku As String
    Dim nomeProduto As String
    Dim categoria As String

    On Error GoTo TratarErro

    Set wsForm = ThisWorkbook.Worksheets("CADASTRO_PRODUTO")
    Set wsBase = ThisWorkbook.Worksheets("BASE_PRODUTOS")
    Set tbl = wsBase.ListObjects("tbProdutos")

    sku = Trim$(CStr(wsForm.Range("B3").Value))
    nomeProduto = Trim$(CStr(wsForm.Range("B4").Value))
    categoria = Trim$(CStr(wsForm.Range("B5").Value))

    If sku = "" Or nomeProduto = "" Or categoria = "" Then
        MsgBox "Preencha SKU, produto e categoria.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(wsForm.Range("B6").Value) _
    Or Not IsNumeric(wsForm.Range("B7").Value) _
    Or Not IsNumeric(wsForm.Range("B8").Value) _
    Or Not IsNumeric(wsForm.Range("B9").Value) Then

        MsgBox "Preencha corretamente custos, precos e estoques.", _
               vbExclamation
        Exit Sub
    End If

    If CDbl(wsForm.Range("B6").Value) < 0 _
    Or CDbl(wsForm.Range("B7").Value) < 0 _
    Or CDbl(wsForm.Range("B8").Value) < 0 _
    Or CDbl(wsForm.Range("B9").Value) < 0 Then

        MsgBox "Os valores nao podem ser negativos.", vbExclamation
        Exit Sub
    End If

    If CDbl(wsForm.Range("B8").Value) <> _
       Fix(CDbl(wsForm.Range("B8").Value)) _
    Or CDbl(wsForm.Range("B9").Value) <> _
       Fix(CDbl(wsForm.Range("B9").Value)) Then

        MsgBox "Os estoques precisam ser numeros inteiros.", _
               vbExclamation
        Exit Sub
    End If

    If Not tbl.DataBodyRange Is Nothing Then
        If Application.CountIf( _
            tbl.ListColumns("SKU").DataBodyRange, sku) > 0 Then

            MsgBox "Este SKU ja esta cadastrado.", vbExclamation
            Exit Sub

        End If
    End If

    If tbl.DataBodyRange Is Nothing Then
        Set novaLinha = tbl.ListRows.Add

    ElseIf Application.WorksheetFunction.CountA( _
        tbl.DataBodyRange.Rows(1)) = 0 Then

        Set novaLinha = tbl.ListRows(1)

    Else
        Set novaLinha = tbl.ListRows.Add
    End If

    With novaLinha.Range

        .Cells(1, tbl.ListColumns("SKU").Index).Value = sku

        .Cells(1, tbl.ListColumns("Produto").Index).Value = _
            nomeProduto

        .Cells(1, tbl.ListColumns("Categoria").Index).Value = _
            categoria

        .Cells(1, tbl.ListColumns("CustoUnitario").Index).Value = _
            CDbl(wsForm.Range("B6").Value)

        .Cells(1, tbl.ListColumns("PrecoVenda").Index).Value = _
            CDbl(wsForm.Range("B7").Value)

        .Cells(1, tbl.ListColumns("EstoqueInicial").Index).Value = _
            CLng(wsForm.Range("B8").Value)

        .Cells(1, tbl.ListColumns("EstoqueMinimo").Index).Value = _
            CLng(wsForm.Range("B9").Value)

        .Cells(1, tbl.ListColumns("DataCadastro").Index).Value = Date

    End With

    wsForm.Range("B3:B9").ClearContents
    ThisWorkbook.Save

    MsgBox "Produto cadastrado com sucesso!", vbInformation
    Exit Sub

TratarErro:
    MsgBox "Erro ao cadastrar produto: " & Err.Description, vbCritical

End Sub


' =========================================================
' 2. REGISTRAR VENDA
' =========================================================

Public Sub RegistrarVenda()

    Dim wsForm As Worksheet
    Dim wsProdutos As Worksheet
    Dim wsVendas As Worksheet

    Dim tblProdutos As ListObject
    Dim tblVendas As ListObject
    Dim novaLinha As ListRow

    Dim sku As String
    Dim nomeProduto As String
    Dim pagamento As String
    Dim vendaID As String

    Dim dataVenda As Date
    Dim quantidade As Long
    Dim desconto As Double
    Dim preco As Double
    Dim custo As Double
    Dim subtotal As Double
    Dim total As Double
    Dim estoqueInicial As Double
    Dim vendidoAntes As Double

    Dim indiceProduto As Variant

    On Error GoTo TratarErro

    Set wsForm = ThisWorkbook.Worksheets("LANCAR_VENDA")
    Set wsProdutos = ThisWorkbook.Worksheets("BASE_PRODUTOS")
    Set wsVendas = ThisWorkbook.Worksheets("BASE_VENDAS")

    Set tblProdutos = wsProdutos.ListObjects("tbProdutos")
    Set tblVendas = wsVendas.ListObjects("tbVendas")

    If Trim$(CStr(wsForm.Range("B4").Value)) = "" Then
        dataVenda = Date

    ElseIf IsDate(wsForm.Range("B4").Value) Then
        dataVenda = CDate(wsForm.Range("B4").Value)

    Else
        MsgBox "Informe uma data valida.", vbExclamation
        Exit Sub
    End If

    sku = Trim$(CStr(wsForm.Range("B5").Value))
    pagamento = Trim$(CStr(wsForm.Range("B8").Value))

    If sku = "" Or pagamento = "" Then
        MsgBox "Informe o SKU e a forma de pagamento.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(wsForm.Range("B6").Value) Then
        MsgBox "Informe uma quantidade valida.", vbExclamation
        Exit Sub
    End If

    If CDbl(wsForm.Range("B6").Value) < 1 _
    Or CDbl(wsForm.Range("B6").Value) <> _
       Fix(CDbl(wsForm.Range("B6").Value)) Then

        MsgBox "A quantidade deve ser um inteiro positivo.", _
               vbExclamation
        Exit Sub
    End If

    quantidade = CLng(wsForm.Range("B6").Value)

    If Trim$(CStr(wsForm.Range("B7").Value)) = "" Then
        desconto = 0

    ElseIf IsNumeric(wsForm.Range("B7").Value) Then
        desconto = CDbl(wsForm.Range("B7").Value)

    Else
        MsgBox "Informe um desconto valido.", vbExclamation
        Exit Sub
    End If

    If desconto < 0 Then
        MsgBox "O desconto nao pode ser negativo.", vbExclamation
        Exit Sub
    End If

    If tblProdutos.DataBodyRange Is Nothing Then
        MsgBox "Cadastre um produto antes de registrar vendas.", _
               vbExclamation
        Exit Sub
    End If

    indiceProduto = Application.Match( _
        sku, tblProdutos.ListColumns("SKU").DataBodyRange, 0)

    If IsError(indiceProduto) Then
        MsgBox "SKU nao encontrado no cadastro.", vbExclamation
        Exit Sub
    End If

    With tblProdutos

        nomeProduto = .ListColumns("Produto").DataBodyRange.Cells( _
            CLng(indiceProduto), 1).Value

        preco = CDbl(.ListColumns("PrecoVenda").DataBodyRange.Cells( _
            CLng(indiceProduto), 1).Value)

        custo = CDbl(.ListColumns("CustoUnitario").DataBodyRange.Cells( _
            CLng(indiceProduto), 1).Value)

        estoqueInicial = CDbl( _
            .ListColumns("EstoqueInicial").DataBodyRange.Cells( _
            CLng(indiceProduto), 1).Value)

    End With

    ' Saldo estimado: estoque inicial menos vendas registradas.
    vendidoAntes = 0

    If Not tblVendas.DataBodyRange Is Nothing Then
        vendidoAntes = Application.SumIf( _
            tblVendas.ListColumns("SKU").DataBodyRange, _
            sku, _
            tblVendas.ListColumns("Quantidade").DataBodyRange)
    End If

    If quantidade > estoqueInicial - vendidoAntes Then
        MsgBox "Estoque insuficiente para esta venda.", vbExclamation
        Exit Sub
    End If

    subtotal = quantidade * preco

    If desconto > subtotal Then
        MsgBox "O desconto nao pode superar o subtotal.", vbExclamation
        Exit Sub
    End If

    total = subtotal - desconto

    ' Reutilizar a primeira linha se ela estiver vazia.
    If tblVendas.DataBodyRange Is Nothing Then
        Set novaLinha = tblVendas.ListRows.Add

    ElseIf Len(Trim$(CStr( _
        tblVendas.ListColumns("VendaID").DataBodyRange.Cells(1, 1).Value))) = 0 _
    And Len(Trim$(CStr( _
        tblVendas.ListColumns("SKU").DataBodyRange.Cells(1, 1).Value))) = 0 Then

        Set novaLinha = tblVendas.ListRows(1)

    Else
        Set novaLinha = tblVendas.ListRows.Add
    End If

    vendaID = "VEND-" & Format(Now, "yyyymmdd-hhnnss") _
        & "-" & Format(tblVendas.ListRows.Count, "0000")

    With novaLinha.Range

        .Cells(1, tblVendas.ListColumns("Data").Index).Value = dataVenda
        .Cells(1, tblVendas.ListColumns("VendaID").Index).Value = vendaID
        .Cells(1, tblVendas.ListColumns("SKU").Index).Value = sku
        .Cells(1, tblVendas.ListColumns("Produto").Index).Value = nomeProduto

        .Cells(1, tblVendas.ListColumns("Quantidade").Index).Value = quantidade
        .Cells(1, tblVendas.ListColumns("PrecoUnitario").Index).Value = preco
        .Cells(1, tblVendas.ListColumns("CustoUnitario").Index).Value = custo
        .Cells(1, tblVendas.ListColumns("Desconto").Index).Value = desconto
        .Cells(1, tblVendas.ListColumns("TotalVenda").Index).Value = total
        .Cells(1, tblVendas.ListColumns("FormaPagamento").Index).Value = pagamento

    End With

    wsForm.Range("B5:B8").ClearContents
    wsForm.Range("B4").Value = Date
    ThisWorkbook.Save

    MsgBox "Venda registrada com sucesso!" & vbCrLf & _
           "Identificador: " & vendaID & vbCrLf & _
           "Total: R$ " & Format(total, "#,##0.00"), vbInformation

    Exit Sub

TratarErro:
    MsgBox "Erro ao registrar venda: " & Err.Description, vbCritical

End Sub


' =========================================================
' 3. GERAR RELATORIO COM PYTHON
' =========================================================

Public Sub GerarRelatorioPython()

    Dim ws As Worksheet
    Dim dataInicial As Date
    Dim dataFinal As Date

    Dim pastaSistema As String
    Dim caminhoExe As String
    Dim pastaRelatorios As String
    Dim caminhoPDF As String

    Dim comando As String
    Dim shell As Object
    Dim codigoSaida As Long

    On Error GoTo TratarErro

    Set ws = ThisWorkbook.Worksheets("RELATORIO")

    If ThisWorkbook.Path = "" Then
        MsgBox "Salve a planilha antes de gerar o relatorio.", _
               vbExclamation
        Exit Sub
    End If

    If Not IsDate(ws.Range("B5").Value) _
    Or Not IsDate(ws.Range("B6").Value) Then

        MsgBox "Preencha as datas inicial e final.", vbExclamation
        Exit Sub
    End If

    dataInicial = CDate(ws.Range("B5").Value)
    dataFinal = CDate(ws.Range("B6").Value)

    If dataFinal < dataInicial Then
        MsgBox "A data final nao pode ser anterior a inicial.", _
               vbExclamation
        Exit Sub
    End If

    pastaSistema = ThisWorkbook.Path

    caminhoExe = pastaSistema & Application.PathSeparator & _
                 "RelatorioFacilSistema.exe"

    If Dir(caminhoExe, vbNormal) = "" Then
        MsgBox "O executavel RelatorioFacilSistema.exe nao foi encontrado." _
               & vbCrLf & "Ele deve ficar na mesma pasta da planilha.", _
               vbExclamation
        Exit Sub
    End If

    pastaRelatorios = pastaSistema & _
                      Application.PathSeparator & "Relatorios"

    If Dir(pastaRelatorios, vbDirectory) = "" Then
        MkDir pastaRelatorios
    End If

    caminhoPDF = pastaRelatorios & _
        Application.PathSeparator & _
        "RelatorioFacil_" & _
        Format(dataInicial, "yyyymmdd") & "_" & _
        Format(dataFinal, "yyyymmdd") & "_" & _
        Format(Now, "hhnnss") & ".pdf"

    ' Atualizar o ranking conforme o periodo selecionado.
    Call AtualizarTop10Periodo

    ' Salvar os dados antes de chamar o Python.
    ThisWorkbook.Save

    comando = Chr(34) & caminhoExe & Chr(34) & _
        " --excel " & Chr(34) & ThisWorkbook.FullName & Chr(34) & _
        " --inicio " & Chr(34) & Format(dataInicial, "yyyy-mm-dd") & Chr(34) & _
        " --fim " & Chr(34) & Format(dataFinal, "yyyy-mm-dd") & Chr(34) & _
        " --saida " & Chr(34) & caminhoPDF & Chr(34)

    ws.Range("B10").Value = "Gerando relatorio..."

    Set shell = CreateObject("WScript.Shell")
    codigoSaida = shell.Run(comando, 0, True)

    If codigoSaida <> 0 Then

        ws.Range("B10").Value = "Erro ao gerar relatorio."

        MsgBox "O sistema Python retornou um erro." & vbCrLf & _
               "Codigo: " & codigoSaida, vbCritical
        Exit Sub
    End If

    If Dir(caminhoPDF, vbNormal) = "" Then

        ws.Range("B10").Value = "PDF nao encontrado."

        MsgBox "O programa terminou, mas o PDF nao foi encontrado.", _
               vbCritical
        Exit Sub
    End If

    If FileLen(caminhoPDF) = 0 Then

        ws.Range("B10").Value = "PDF vazio."

        MsgBox "O PDF foi criado, mas esta vazio.", vbCritical
        Exit Sub
    End If

    ws.Range("B10").Value = _
        "Relatorio gerado em " & Format(Now, "dd/mm/yyyy hh:nn")

    MsgBox "Relatorio gerado com sucesso!" & vbCrLf & vbCrLf & _
           "Arquivo:" & vbCrLf & caminhoPDF, vbInformation

    Exit Sub

TratarErro:
    MsgBox "Erro ao gerar relatorio: " & Err.Description, vbCritical

End Sub


Public Sub AtualizarTop10Periodo()

    Dim wsRel As Worksheet
    Dim wsVendas As Worksheet
    Dim tblVendas As ListObject

    Dim dict As Object
    Dim chaves As Variant
    Dim info As Variant
    Dim dados() As Variant

    Dim i As Long, j As Long, k As Long
    Dim n As Long, limite As Long
    Dim linha As Long

    Dim dataInicial As Date
    Dim dataFinal As Date
    Dim dataVenda As Date

    Dim sku As String
    Dim nomeProduto As String

    Dim quantidade As Double
    Dim preco As Double
    Dim custo As Double
    Dim desconto As Double
    Dim total As Double

    Dim temporario As Variant

    On Error GoTo TratarErro

    Set wsRel = ThisWorkbook.Worksheets("RELATORIO")
    Set wsVendas = ThisWorkbook.Worksheets("BASE_VENDAS")
    Set tblVendas = wsVendas.ListObjects("tbVendas")

    ' Validar o periodo informado
    If Not IsDate(wsRel.Range("B5").Value) _
    Or Not IsDate(wsRel.Range("B6").Value) Then

        MsgBox "Preencha as datas inicial e final.", vbExclamation
        Exit Sub

    End If

    dataInicial = DateValue(wsRel.Range("B5").Value)
    dataFinal = DateValue(wsRel.Range("B6").Value)

    If dataFinal < dataInicial Then
        MsgBox "A data final nao pode ser anterior a inicial.", _
               vbExclamation
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    ' Ler as vendas do periodo e agrupar por SKU
    If Not tblVendas.DataBodyRange Is Nothing Then

        For i = 1 To tblVendas.ListRows.Count

            If IsDate(tblVendas.ListColumns("Data") _
                .DataBodyRange.Cells(i, 1).Value) Then

                dataVenda = DateValue( _
                    tblVendas.ListColumns("Data") _
                    .DataBodyRange.Cells(i, 1).Value)

                If dataVenda >= dataInicial _
                And dataVenda <= dataFinal Then

                    sku = Trim$(CStr( _
                        tblVendas.ListColumns("SKU") _
                        .DataBodyRange.Cells(i, 1).Value))

                    If sku <> "" Then

                        nomeProduto = Trim$(CStr( _
                            tblVendas.ListColumns("Produto") _
                            .DataBodyRange.Cells(i, 1).Value))

                        quantidade = Val( _
                            tblVendas.ListColumns("Quantidade") _
                            .DataBodyRange.Cells(i, 1).Value)

                        preco = Val( _
                            tblVendas.ListColumns("PrecoUnitario") _
                            .DataBodyRange.Cells(i, 1).Value)

                        custo = Val( _
                            tblVendas.ListColumns("CustoUnitario") _
                            .DataBodyRange.Cells(i, 1).Value)

                        desconto = Val( _
                            tblVendas.ListColumns("Desconto") _
                            .DataBodyRange.Cells(i, 1).Value)

                        total = quantidade * preco - desconto

                        If dict.Exists(sku) Then
                            info = dict(sku)
                        Else
                            If nomeProduto = "" Then
                                nomeProduto = sku
                            End If

                            info = Array(nomeProduto, 0#, 0#, 0#)
                        End If

                        info(1) = CDbl(info(1)) + quantidade
                        info(2) = CDbl(info(2)) + total
                        info(3) = CDbl(info(3)) + quantidade * custo

                        dict(sku) = info

                    End If

                End If

            End If

        Next i

    End If

    ' Limpar resultados anteriores, preservando os cabecalhos
    wsRel.Range("A32:D41").ClearContents

    If dict.Count = 0 Then

        wsRel.Range("A32").Value = _
            "Nenhuma venda encontrada no periodo."

        MsgBox "Nao existem vendas no periodo selecionado.", _
               vbInformation

        Exit Sub

    End If

    ' Montar a matriz do ranking
    n = dict.Count
    ReDim dados(1 To n, 1 To 4)

    chaves = dict.Keys

    For i = 0 To n - 1

        info = dict(chaves(i))

        dados(i + 1, 1) = info(0)
        dados(i + 1, 2) = info(1)
        dados(i + 1, 3) = info(2)
        dados(i + 1, 4) = info(2) - info(3)

    Next i

    ' Ordenar por quantidade decrescente.
    ' Em caso de empate, maior faturamento primeiro.
    For i = 1 To n - 1

        For j = i + 1 To n

            If dados(j, 2) > dados(i, 2) _
            Or (dados(j, 2) = dados(i, 2) _
                And dados(j, 3) > dados(i, 3)) Then

                For k = 1 To 4
                    temporario = dados(i, k)
                    dados(i, k) = dados(j, k)
                    dados(j, k) = temporario
                Next k

            End If

        Next j

    Next i

    ' Escrever os dez primeiros resultados na planilha
    limite = n

    If limite > 10 Then limite = 10

    For i = 1 To limite

        linha = 31 + i

        wsRel.Cells(linha, 1).Value = dados(i, 1)
        wsRel.Cells(linha, 2).Value = dados(i, 2)
        wsRel.Cells(linha, 3).Value = dados(i, 3)
        wsRel.Cells(linha, 4).Value = dados(i, 4)

        wsRel.Cells(linha, 2).NumberFormat = "0"
        wsRel.Cells(linha, 3).NumberFormat = """R$"" #,##0.00"
        wsRel.Cells(linha, 4).NumberFormat = """R$"" #,##0.00"

    Next i

    ThisWorkbook.Save

    Exit Sub

TratarErro:

    MsgBox "Erro ao atualizar o Top 10: " & _
           Err.Description, vbCritical

End Sub


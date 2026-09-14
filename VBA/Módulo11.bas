Attribute VB_Name = "Módulo11"
Private ZonasAstra As Object
Private MargemSeguraAstra As Double
Private ModoLoteAstra As Boolean

Sub AjustarECentralizar()

    Dim sr As ShapeRange
    Dim largura As Double
    Dim altura As Double
    Dim escala As Double
    Dim margem As Double

    ' Verifica se existe algum objeto selecionado
    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Selecione um objeto primeiro.", vbExclamation, "Assistente Corel"
        Exit Sub
    End If

    Set sr = ActiveSelectionRange

    ' Trabalhar em milímetros
    ActiveDocument.Unit = cdrMillimeter

    ' Usar o centro do objeto como referência
    ActiveDocument.ReferencePoint = cdrCenter

    ' Margem de segurança da página
    margem = 10

    ' Calcula escala pela largura
    escala = (ActivePage.SizeWidth - margem * 2) / sr.SizeWidth

    ' Verifica se pela altura ficaria grande demais
    If sr.SizeHeight * escala > ActivePage.SizeHeight - margem * 2 Then
        escala = (ActivePage.SizeHeight - margem * 2) / sr.SizeHeight
    End If

    largura = sr.SizeWidth * escala
    altura = sr.SizeHeight * escala

    ' Redimensiona proporcionalmente
    sr.SetSize largura, altura

    ' Centraliza exatamente na página
    sr.SetPosition ActivePage.SizeWidth / 2, ActivePage.SizeHeight / 2

End Sub

Sub EnviarPedidoAoAstra()

    Dim pedido As String
    Dim arquivo As Integer
    Dim caminho As String

    pedido = InputBox( _
        "O que você quer fazer no Corel?", _
        "Corel AI" _
    )

    If Trim(pedido) = "" Then Exit Sub

    caminho = "C:\CorelAI\pedido.txt"

    arquivo = FreeFile

    Open caminho For Output As #arquivo
    Print #arquivo, pedido
    Close #arquivo

    MsgBox _
        "Pedido enviado para C:\CorelAI\pedido.txt", _
        vbInformation, _
        "Corel AI"

End Sub

Sub ExecutarComandoDoAstra()

    Dim caminho As String
    Dim arquivo As Integer
    Dim comando As String

    caminho = "C:\CorelAI\comando.txt"

    If Dir(caminho) = "" Then
        MsgBox "O arquivo comando.txt não foi encontrado.", vbExclamation, "Corel AI"
        Exit Sub
    End If

    arquivo = FreeFile

    Open caminho For Input As #arquivo
    Line Input #arquivo, comando
    Close #arquivo

    comando = Trim(UCase(comando))

    Select Case comando

        Case "AJUSTAR"
            AjustarECentralizar

        Case Else
            MsgBox _
                "Comando não reconhecido: " & comando, _
                vbExclamation, _
                "Corel AI"

    End Select

End Sub
Sub ExportarCenaParaAstra()

    Dim s As Shape
    Dim sr As ShapeRange
    Dim arq As Integer
    Dim caminho As String
    Dim i As Long

    caminho = "C:\CorelAI\cena.txt"

    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter

    Set sr = ActivePage.Shapes.All

    arq = FreeFile
    Open caminho For Output As #arq

    Print #arq, "PAGINA"
    Print #arq, "LARGURA_MM=" & NumeroTexto(ActivePage.SizeWidth)
    Print #arq, "ALTURA_MM=" & NumeroTexto(ActivePage.SizeHeight)
    Print #arq, "TOTAL_OBJETOS=" & sr.Count
    Print #arq, ""

    For i = 1 To sr.Count

        Set s = sr(i)

        Print #arq, "OBJETO_" & i
        Print #arq, "ID=" & i
        Print #arq, "ORDEM=" & i
        Print #arq, "TIPO=" & TipoDoObjeto(s)

        If Trim(s.Name) <> "" Then
            Print #arq, "NOME=" & LimparTextoCena(s.Name)
        End If

        Print #arq, "X_MM=" & NumeroTexto(s.CenterX)
        Print #arq, "Y_MM=" & NumeroTexto(s.CenterY)
        Print #arq, "LARGURA_MM=" & NumeroTexto(s.SizeWidth)
        Print #arq, "ALTURA_MM=" & NumeroTexto(s.SizeHeight)

        ExportarInfoTexto s, arq
        ExportarInfoGrupoDetalhado s, arq, 1, CStr(i)
        ExportarInfoPreenchimento s, arq
        ExportarInfoContorno s, arq

        Print #arq, ""

    Next i
    
        Print #arq, ""
        Print #arq, "MAPA_SEMANTICO"
        Print #arq, "ARQUIVO_ROLES=C:\CorelAI\roles.txt"

    Close #arq

    MsgBox _
        "Cena detalhada exportada." & vbCrLf & _
        sr.Count & " objetos encontrados.", _
        vbInformation, _
        "Corel AI"

End Sub


Function TipoDoObjeto(ByVal s As Shape) As String

    Select Case s.Type

        Case cdrTextShape
            TipoDoObjeto = "TEXTO"

        Case cdrBitmapShape
            TipoDoObjeto = "BITMAP"

        Case cdrRectangleShape
            TipoDoObjeto = "RETANGULO"

        Case cdrEllipseShape
            TipoDoObjeto = "ELIPSE"

        Case cdrCurveShape
            TipoDoObjeto = "CURVA"

        Case cdrGroupShape
            TipoDoObjeto = "GRUPO"

        Case Else
            TipoDoObjeto = "OUTRO_" & CStr(s.Type)

    End Select

End Function

Sub ExportarPreviewParaAstra()

    Dim caminho As String
    Dim filtro As ExportFilter

    Dim larguraPagina As Double
    Dim alturaPagina As Double

    Dim larguraPx As Long
    Dim alturaPx As Long

    caminho = "C:\CorelAI\preview.png"

    ActiveDocument.Unit = cdrMillimeter

    larguraPagina = ActivePage.SizeWidth
    alturaPagina = ActivePage.SizeHeight

    ' largura do preview
    larguraPx = 1200

    ' calcula a altura mantendo exatamente o aspect ratio da página
    alturaPx = CLng(larguraPx * (alturaPagina / larguraPagina))

    Set filtro = ActiveDocument.ExportBitmap( _
        caminho, _
        cdrPNG, _
        cdrCurrentPage, _
        cdrRGBColorImage, _
        larguraPx, _
        alturaPx, _
        96, _
        96 _
    )

    filtro.Finish

    MsgBox _
        "Preview exportado corretamente." & vbCrLf & vbCrLf & _
        "Página: " & larguraPagina & " x " & alturaPagina & " mm" & vbCrLf & _
        "Imagem: " & larguraPx & " x " & alturaPx & " px", _
        vbInformation, _
        "Corel AI"

End Sub

Sub PrepararArteParaAstra()

    Dim resposta As VbMsgBoxResult

    ' 1. Mapeia todos os objetos da página
    ExportarCenaParaAstra

    ' 2. Gera imagem visual da página
    ExportarPreviewParaAstra

    ' 3. Pergunta o que o usuário quer fazer
    EnviarPedidoAoAstra

    ' 4. Confirma conclusão
    MsgBox _
        "Pacote preparado para o Astra." & vbCrLf & vbCrLf & _
        "Foram gerados:" & vbCrLf & _
        "• cena.txt" & vbCrLf & _
        "• preview.png" & vbCrLf & _
        "• pedido.txt", _
        vbInformation, _
        "Corel AI"

End Sub

Sub AplicarLayoutArquivo( _
    ByVal caminho As String, _
    Optional ByVal duplicarPagina As Boolean = True)

    Dim arq As Integer
    Dim linha As String
    Dim partes() As String

    Dim novaPagina As Page
    Dim paginaOriginal As Page

    If Dir(caminho) = "" Then
    If Not ValidarLayoutAstra(ActivePage, caminho) Then
    Exit Sub
    End If
        MsgBox _
            "Arquivo não encontrado:" & vbCrLf & caminho, _
            vbExclamation, _
            "Corel AI"
        Exit Sub
    End If
    
    Set ZonasAstra = CreateObject("Scripting.Dictionary")
    
    MargemSeguraAstra = 0

    Set paginaOriginal = ActivePage

    If duplicarPagina Then

    paginaOriginal.Activate
    paginaOriginal.Shapes.All.CreateSelection
    ActiveSelectionRange.Copy

    Set novaPagina = ActiveDocument.AddPages(1)

    novaPagina.Activate
    novaPagina.ActiveLayer.Paste

    Else

    Set novaPagina = ActivePage

    End If

    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter

    arq = FreeFile
    Open caminho For Input As #arq

    Do Until EOF(arq)

        Line Input #arq, linha
        linha = Trim(linha)

        If linha <> "" Then

            partes = Split(linha, "|")

            Select Case UCase(partes(0))
            
                Case "NOCOLLIDE"

                If UBound(partes) >= 4 Then

                NoCollideAstra _
                novaPagina, _
                partes(1), _
                partes(2), _
                partes(3), _
                NumeroAstra(partes(4))

                End If
            
                Case "SAFEAREA"

                If UBound(partes) >= 2 Then

                DefinirSafeAreaAstra _
                novaPagina, _
                NumeroAstra(partes(1)), _
                partes(2)

                End If


                Case "AUTOFIT"

                If UBound(partes) >= 1 Then

                AutoFitAstra _
                novaPagina, _
                partes(1)

                End If
            
                Case "GAP"

                If UBound(partes) >= 5 Then

                GapAstra _
                novaPagina, _
                partes(1), _
                partes(2), _
                partes(3), _
                NumeroAstra(partes(4)), _
                partes(5)

                End If
            
                Case "ZONE"

                If UBound(partes) >= 5 Then

                DefinirZonaAstra _
                partes(1), _
                NumeroAstra(partes(2)), _
                NumeroAstra(partes(3)), _
                NumeroAstra(partes(4)), _
                NumeroAstra(partes(5))

                End If


                Case "PLACEZONE"

                If UBound(partes) >= 5 Then

                PosicionarNaZonaAstra _
                novaPagina, _
                partes(1), _
                partes(2), _
                partes(3), _
                partes(4), _
                NumeroAstra(partes(5))

                End If
            
                Case "KEEPTOGETHER"

                If UBound(partes) >= 3 Then
                ManterJuntosAstra _
                novaPagina, _
                partes(1), _
                NumeroAstra(partes(2)), _
                NumeroAstra(partes(3))
                End If


                Case "ABOVE"

                If UBound(partes) >= 3 Then
                AcimaDeAstra _
                novaPagina, _
                partes(1), _
                partes(2), _
                NumeroAstra(partes(3))
                End If


                Case "DISTRIBUTE"

                If UBound(partes) >= 2 Then
                DistribuirAstra _
                novaPagina, _
                partes(1), _
                partes(2)
                End If
            
                Case "ALIGN"

                If UBound(partes) >= 3 Then
                AlinharAstra _
                novaPagina, _
                partes(1), _
                partes(2), _
                partes(3)
                End If


                Case "BELOW"

                If UBound(partes) >= 3 Then
                AbaixoDeAstra _
                novaPagina, _
                partes(1), _
                partes(2), _
                NumeroAstra(partes(3))
                End If


                Case "RIGHTOF"

                If UBound(partes) >= 3 Then
                DireitaDeAstra _
                novaPagina, _
                partes(1), _
                partes(2), _
                NumeroAstra(partes(3))
                End If


                Case "MARGIN"

                If UBound(partes) >= 3 Then

                MargemSeletorAstra _
                novaPagina, _
                partes(1), _
                partes(2), _
                NumeroAstra(partes(3))

                End If

                Case "PAGE"

                    novaPagina.SetSize _
                        NumeroAstra(partes(1)), _
                        NumeroAstra(partes(2))


                Case "FILLPAGE"

                    PreencherPaginaAstra _
                        novaPagina, _
                        (partes(1))


                Case "OBJECT"

                    AplicarObjetoAstra _
                        novaPagina, _
                        (partes(1)), _
                        NumeroAstra(partes(2)), _
                        NumeroAstra(partes(3)), _
                        NumeroAstra(partes(4))


                Case "SET"

                    AplicarConjuntoAstra _
                        novaPagina, _
                        partes(1), _
                        NumeroAstra(partes(2)), _
                        NumeroAstra(partes(3)), _
                        NumeroAstra(partes(4))


                Case "STRETCH"

                    EsticarObjetoAstra _
                        novaPagina, _
                        (partes(1)), _
                        NumeroAstra(partes(2)), _
                        NumeroAstra(partes(3)), _
                        NumeroAstra(partes(4)), _
                        NumeroAstra(partes(5))


                Case "FITINSIDE"

                AjustarDentroAstra _
                    novaPagina, _
                    partes(1), _
                    partes(2), _
                    NumeroAstra(partes(3))


                Case "SCALE"

                    EscalarObjetoAstra _
                        novaPagina, _
                        (partes(1)), _
                        NumeroAstra(partes(2))

            End Select

        End If

    Loop

    Close #arq
    
    ValidarResultadoAstra novaPagina
    
    AvaliarLayoutSemanticoAstra novaPagina
    
    SalvarAnaliseSemanticaAstra novaPagina
    
    AvaliarColisoesAstra novaPagina
    
    SalvarRelatorioLayoutAstra novaPagina
    

    MsgBox _
        "Layout aplicado em uma nova página." & vbCrLf & vbCrLf & _
        "Arquivo usado:" & vbCrLf & caminho, _
        vbInformation, _
        "Corel AI"

End Sub


Sub AplicarObjetoAstra( _
    ByVal pg As Page, _
    ByVal id As String, _
    ByVal novoX As Double, _
    ByVal novoY As Double, _
    ByVal novaLargura As Double)

    Dim s As Shape

    Set s = ObjetoPorIDAstra(pg, id)

    If s Is Nothing Then Exit Sub

    ActiveDocument.ReferencePoint = cdrCenter

    s.SetSize novaLargura, 0
    s.SetPosition novoX, novoY

End Sub

Sub AplicarConjuntoAstra( _
    ByVal pg As Page, _
    ByVal listaIDs As String, _
    ByVal novoX As Double, _
    ByVal novoY As Double, _
    ByVal novaLargura As Double)

    Dim ids() As String
    Dim i As Long

    Dim s As Shape
    Dim conjunto As New ShapeRange
    
    listaIDs = ResolverSeletorAstra(listaIDs)

    If Trim(listaIDs) = "" Then Exit Sub

    ids = Split(listaIDs, ",")

    For i = LBound(ids) To UBound(ids)

        Set s = ObjetoPorIDAstra(pg, Trim(ids(i)))

        If Not s Is Nothing Then
            conjunto.Add s
        End If

    Next i

    If conjunto.Count = 0 Then Exit Sub

    ActiveDocument.ReferencePoint = cdrCenter

    conjunto.SetSize novaLargura, 0
    conjunto.SetPosition novoX, novoY

End Sub


Function NumeroAstra(ByVal valor As String) As Double

    valor = Trim(valor)

    ' Val entende ponto decimal independentemente
    ' da configuração regional do Windows
    NumeroAstra = Val(valor)

End Function
Sub PreencherPaginaAstra(ByVal pg As Page, ByVal id As String)

    Dim s As Shape

    Set s = ObjetoPorIDAstra(pg, id)

    If s Is Nothing Then Exit Sub

    ActiveDocument.ReferencePoint = cdrCenter

    s.SetSize pg.SizeWidth, pg.SizeHeight
    s.SetPosition pg.SizeWidth / 2, pg.SizeHeight / 2

End Sub

Sub ExportarPreviewResultado()

    Dim caminho As String
    Dim filtro As ExportFilter
    Dim larguraPagina As Double
    Dim alturaPagina As Double
    Dim larguraPx As Long
    Dim alturaPx As Long

    caminho = "C:\CorelAI\preview_resultado.png"

    ActiveDocument.Unit = cdrMillimeter

    larguraPagina = ActivePage.SizeWidth
    alturaPagina = ActivePage.SizeHeight

    ' Preview horizontal com boa resolução
    larguraPx = 2000

    ' Mantém exatamente o aspect ratio da página
    alturaPx = CLng(larguraPx * (alturaPagina / larguraPagina))

    Set filtro = ActiveDocument.ExportBitmap( _
        caminho, _
        cdrPNG, _
        cdrCurrentPage, _
        cdrRGBColorImage, _
        larguraPx, _
        alturaPx, _
        96, _
        96 _
    )

    filtro.Finish

    MsgBox _
        "Resultado exportado para análise do Astra." & vbCrLf & vbCrLf & _
        "Página: " & larguraPagina & " x " & alturaPagina & " mm" & vbCrLf & _
        "Preview: " & larguraPx & " x " & alturaPx & " px", _
        vbInformation, _
        "Corel AI"

End Sub
Sub EsticarObjetoAstra( _
    ByVal pg As Page, _
    ByVal id As String, _
    ByVal novoX As Double, _
    ByVal novoY As Double, _
    ByVal novaLargura As Double, _
    ByVal novaAltura As Double)

    Dim s As Shape

    Set s = ObjetoPorIDAstra(pg, id)

    If s Is Nothing Then Exit Sub

    ActiveDocument.ReferencePoint = cdrCenter

    s.SetSize novaLargura, novaAltura
    s.SetPosition novoX, novoY

End Sub
Sub AjustarDentroAstra( _
    ByVal pg As Page, _
    ByVal listaIDs As String, _
    ByVal idContainer As String, _
    ByVal fator As Double)

    Dim ids() As String
    Dim i As Long

    Dim s As Shape
    Dim container As Shape
    Dim conjunto As New ShapeRange

    Dim larguraMax As Double
    Dim alturaMax As Double

    Dim escalaL As Double
    Dim escalaA As Double
    Dim escala As Double

    Set container = ObjetoPorIDAstra(pg, idContainer)

    If container Is Nothing Then Exit Sub
    
    listaIDs = ResolverSeletorAstra(listaIDs)

    If Trim(listaIDs) = "" Then Exit Sub

    ids = Split(listaIDs, ",")

    For i = LBound(ids) To UBound(ids)

        Set s = ObjetoPorIDAstra(pg, Trim(ids(i)))

        If Not s Is Nothing Then
            conjunto.Add s
        End If

    Next i

    If conjunto.Count = 0 Then Exit Sub

    larguraMax = container.SizeWidth * fator
    alturaMax = container.SizeHeight * fator

    escalaL = larguraMax / conjunto.SizeWidth
    escalaA = alturaMax / conjunto.SizeHeight

    If escalaL < escalaA Then
        escala = escalaL
    Else
        escala = escalaA
    End If

    ActiveDocument.ReferencePoint = cdrCenter

    conjunto.SetSize _
        conjunto.SizeWidth * escala, _
        conjunto.SizeHeight * escala

    conjunto.SetPosition _
        container.CenterX, _
        container.CenterY

End Sub

Sub EscalarObjetoAstra( _
    ByVal pg As Page, _
    ByVal id As String, _
    ByVal fator As Double)

    Dim s As Shape
    Dim x As Double
    Dim y As Double

    Set s = ObjetoPorIDAstra(pg, id)

    If s Is Nothing Then Exit Sub

    x = s.CenterX
    y = s.CenterY

    ActiveDocument.ReferencePoint = cdrCenter

    s.SetSize _
        s.SizeWidth * fator, _
        s.SizeHeight * fator

    s.SetPosition x, y

End Sub

Sub PrepararRevisaoAstra()

    ExportarCenaParaAstra
    ExportarPreviewResultado

    MsgBox _
        "Revisão preparada." & vbCrLf & vbCrLf & _
        "O Astra já pode analisar:" & vbCrLf & _
        "• cena.txt" & vbCrLf & _
        "• preview_resultado.png" & vbCrLf & _
        "• pedido.txt", _
        vbInformation, _
        "Corel AI"

End Sub

Sub AplicarLayoutEscolhido()

    Dim escolha As String
    Dim caminho As String

    escolha = InputBox( _
        "Qual layout deseja aplicar?" & vbCrLf & vbCrLf & _
        "1 = layout.txt" & vbCrLf & _
        "2 = layout_v2.txt", _
        "Corel AI", _
        "2" _
    )

    If Trim(escolha) = "" Then Exit Sub

    Select Case Trim(escolha)

        Case "1"
            caminho = "C:\CorelAI\layout.txt"

        Case "2"
            caminho = "C:\CorelAI\layout_v2.txt"

        Case Else
            MsgBox "Escolha inválida.", vbExclamation, "Corel AI"
            Exit Sub

    End Select

    AplicarLayoutArquivo caminho

End Sub

Sub AplicarUltimoLayout()

    Dim caminho As String
    Dim versao As Long
    Dim proximaVersao As Long
    Dim teste As String

    ' Começa pelo layout base
    caminho = "C:\CorelAI\layout.txt"

    If Dir(caminho) = "" Then
        MsgBox "layout.txt não encontrado.", vbExclamation, "Corel AI"
        Exit Sub
    End If

    versao = 1
    proximaVersao = 2

    Do

        teste = "C:\CorelAI\layout_v" & proximaVersao & ".txt"

        If Dir(teste) <> "" Then

            caminho = teste
            versao = proximaVersao
            proximaVersao = proximaVersao + 1

        Else
            Exit Do
        End If

    Loop

    If versao = 1 Then

        MsgBox _
            "Aplicando layout.txt", _
            vbInformation, _
            "Corel AI"

    Else

        MsgBox _
            "Aplicando layout_v" & versao & ".txt", _
            vbInformation, _
            "Corel AI"

    End If

    AplicarLayoutArquivo caminho

End Sub

Sub PrepararNovaRevisao()

    ' Exporta o resultado visual atual
    ExportarPreviewResultado

    MsgBox _
        "Resultado preparado para nova revisão." & vbCrLf & vbCrLf & _
        "No Codex, peça:" & vbCrLf & vbCrLf & _
        """Leia INSTRUCOES_REVISAO.txt e faça uma nova revisão.""", _
        vbInformation, _
        "Corel AI"

End Sub

Sub ExportarInfoTexto(ByVal s As Shape, ByVal arq As Integer)

    If s.Type <> cdrTextShape Then Exit Sub

    On Error Resume Next

    Print #arq, "TEXTO=" & LimparTextoCena(s.Text.Story.Text)
    Print #arq, "FONTE=" & s.Text.Story.Font
    Print #arq, "TAMANHO_PT=" & NumeroTexto(s.Text.Story.Size)

    On Error GoTo 0

End Sub


Sub ExportarInfoGrupo(ByVal s As Shape, ByVal arq As Integer)

    If s.Type <> cdrGroupShape Then Exit Sub

    On Error Resume Next

    Print #arq, "FILHOS=" & s.Shapes.Count

    On Error GoTo 0

End Sub


Sub ExportarInfoPreenchimento(ByVal s As Shape, ByVal arq As Integer)

    On Error Resume Next

    If s.Fill.Type = cdrUniformFill Then

        Print #arq, _
            "PREENCHIMENTO_CMYK=" & _
            s.Fill.UniformColor.CMYKCyan & "," & _
            s.Fill.UniformColor.CMYKMagenta & "," & _
            s.Fill.UniformColor.CMYKYellow & "," & _
            s.Fill.UniformColor.CMYKBlack

    Else

        Print #arq, "PREENCHIMENTO=OUTRO"

    End If

    On Error GoTo 0

End Sub


Sub ExportarInfoContorno(ByVal s As Shape, ByVal arq As Integer)

    On Error Resume Next

    If s.Outline.Type <> cdrNoOutline Then

        Print #arq, "CONTORNO=SIM"
        Print #arq, "CONTORNO_MM=" & NumeroTexto(s.Outline.Width)

    Else

        Print #arq, "CONTORNO=NAO"

    End If

    On Error GoTo 0

End Sub


Function NumeroTexto(ByVal valor As Double) As String

    NumeroTexto = Replace(Format(valor, "0.00"), ",", ".")

End Function


Function LimparTextoCena(ByVal texto As String) As String

    texto = Replace(texto, vbCr, " ")
    texto = Replace(texto, vbLf, " ")
    texto = Replace(texto, "|", "/")

    LimparTextoCena = Trim(texto)

End Function

Sub ExportarInfoGrupoDetalhado( _
    ByVal s As Shape, _
    ByVal arq As Integer, _
    ByVal nivel As Long, _
    ByVal idPai As String)

    Dim filho As Shape
    Dim i As Long
    Dim prefixo As String
    Dim idFilho As String

    If s.Type <> cdrGroupShape Then Exit Sub

    On Error Resume Next

    Print #arq, "FILHOS=" & s.Shapes.Count

    For i = 1 To s.Shapes.Count

        Set filho = s.Shapes(i)

        prefixo = String(nivel * 2, " ")
        idFilho = idPai & "." & CStr(i)

        Print #arq, prefixo & "FILHO_" & i
        Print #arq, prefixo & "ID=" & idFilho
        Print #arq, prefixo & "NIVEL=" & nivel
        Print #arq, prefixo & "TIPO=" & TipoDoObjeto(filho)

        If Trim(filho.Name) <> "" Then
            Print #arq, prefixo & "NOME=" & LimparTextoCena(filho.Name)
        End If

        Print #arq, prefixo & "X_MM=" & NumeroTexto(filho.CenterX)
        Print #arq, prefixo & "Y_MM=" & NumeroTexto(filho.CenterY)
        Print #arq, prefixo & "LARGURA_MM=" & NumeroTexto(filho.SizeWidth)
        Print #arq, prefixo & "ALTURA_MM=" & NumeroTexto(filho.SizeHeight)

        If filho.Type = cdrTextShape Then
            ExportarInfoTextoIndentado filho, arq, prefixo
        End If

        ExportarInfoPreenchimentoIndentado filho, arq, prefixo
        ExportarInfoContornoIndentado filho, arq, prefixo

        If filho.Type = cdrGroupShape Then

            ExportarInfoGrupoDetalhado _
                filho, _
                arq, _
                nivel + 1, _
                idFilho

        End If

        Print #arq, ""

    Next i

    On Error GoTo 0

End Sub

Sub ExportarInfoTextoIndentado(ByVal s As Shape, ByVal arq As Integer, ByVal prefixo As String)

    On Error Resume Next

    Print #arq, prefixo & "TEXTO=" & LimparTextoCena(s.Text.Story.Text)
    Print #arq, prefixo & "FONTE=" & s.Text.Story.Font
    Print #arq, prefixo & "TAMANHO_PT=" & NumeroTexto(s.Text.Story.Size)

    On Error GoTo 0

End Sub


Sub ExportarInfoPreenchimentoIndentado(ByVal s As Shape, ByVal arq As Integer, ByVal prefixo As String)

    On Error Resume Next

    If s.Fill.Type = cdrUniformFill Then

        Print #arq, prefixo & _
            "PREENCHIMENTO_CMYK=" & _
            s.Fill.UniformColor.CMYKCyan & "," & _
            s.Fill.UniformColor.CMYKMagenta & "," & _
            s.Fill.UniformColor.CMYKYellow & "," & _
            s.Fill.UniformColor.CMYKBlack

    Else

        Print #arq, prefixo & "PREENCHIMENTO=OUTRO"

    End If

    On Error GoTo 0

End Sub


Sub ExportarInfoContornoIndentado(ByVal s As Shape, ByVal arq As Integer, ByVal prefixo As String)

    On Error Resume Next

    If s.Outline.Type <> cdrNoOutline Then
        Print #arq, prefixo & "CONTORNO=SIM"
        Print #arq, prefixo & "CONTORNO_MM=" & NumeroTexto(s.Outline.Width)
    Else
        Print #arq, prefixo & "CONTORNO=NAO"
    End If

    On Error GoTo 0

End Sub

Function ObjetoPorIDAstra(ByVal pg As Page, ByVal idAstra As String) As Shape

    If ModoLoteAstra Then
    Set ObjetoPorIDAstra = ObjetoPorIDLoteAstra(pg, idAstra)
    Exit Function
    End If

    Dim partes() As String
    Dim atual As Shape
    Dim todos As ShapeRange
    Dim i As Long
    Dim indice As Long
    
    If Left(Trim(idAstra), 1) = "@" Then

    idAstra = PrimeiroIDAstra(idAstra)

    If Trim(idAstra) = "" Then
        Set ObjetoPorIDAstra = Nothing
        Exit Function
    End If

    End If

    partes = Split(Trim(idAstra), ".")

    If Not IsNumeric(partes(0)) Then
    Set ObjetoPorIDAstra = Nothing
    Exit Function
    End If


    Set todos = pg.Shapes.All

    indice = CLng(partes(0))

    If indice < 1 Or indice > todos.Count Then
        Set ObjetoPorIDAstra = Nothing
        Exit Function
    End If

    Set atual = todos(indice)

    For i = 1 To UBound(partes)

        If atual.Type <> cdrGroupShape Then
            Set ObjetoPorIDAstra = Nothing
            Exit Function
        End If

        indice = CLng(partes(i))

        If indice < 1 Or indice > atual.Shapes.Count Then
            Set ObjetoPorIDAstra = Nothing
            Exit Function
        End If

        Set atual = atual.Shapes(indice)

    Next i

    Set ObjetoPorIDAstra = atual

End Function

Sub AlinharAstra( _
    ByVal pg As Page, _
    ByVal listaIDs As String, _
    ByVal modo As String, _
    ByVal idReferencia As String)

    Dim ids() As String
    Dim i As Long
    Dim s As Shape
    Dim ref As Shape

    Set ref = ObjetoPorIDAstra(pg, idReferencia)
    If ref Is Nothing Then Exit Sub
    
    listaIDs = ResolverSeletorAstra(listaIDs)

    If Trim(listaIDs) = "" Then Exit Sub

    ids = Split(listaIDs, ",")

    For i = LBound(ids) To UBound(ids)

        Set s = ObjetoPorIDAstra(pg, Trim(ids(i)))
        If s Is Nothing Then GoTo Proximo

        Select Case UCase(Trim(modo))

            Case "CENTER_X"
                s.SetPosition ref.CenterX, s.CenterY

            Case "CENTER_Y"
                s.SetPosition s.CenterX, ref.CenterY

            Case "LEFT"
                s.SetPosition _
                    ref.LeftX + s.SizeWidth / 2, _
                    s.CenterY

            Case "RIGHT"
                s.SetPosition _
                    ref.RightX - s.SizeWidth / 2, _
                    s.CenterY

            Case "TOP"
                s.SetPosition _
                    s.CenterX, _
                    ref.TopY - s.SizeHeight / 2

            Case "BOTTOM"
                s.SetPosition _
                    s.CenterX, _
                    ref.BottomY + s.SizeHeight / 2

        End Select

Proximo:
    Next i

End Sub

Sub AbaixoDeAstra( _
    ByVal pg As Page, _
    ByVal idObjeto As String, _
    ByVal idReferencia As String, _
    ByVal distancia As Double)

    Dim s As Shape
    Dim ref As Shape
    Dim novoY As Double

    Set s = ObjetoPorIDAstra(pg, idObjeto)
    Set ref = ObjetoPorIDAstra(pg, idReferencia)

    If s Is Nothing Then Exit Sub
    If ref Is Nothing Then Exit Sub

    novoY = ref.BottomY - distancia - (s.SizeHeight / 2)

    s.SetPosition s.CenterX, novoY

End Sub

Sub DireitaDeAstra( _
    ByVal pg As Page, _
    ByVal idObjeto As String, _
    ByVal idReferencia As String, _
    ByVal distancia As Double)

    Dim s As Shape
    Dim ref As Shape
    Dim novoX As Double

    Set s = ObjetoPorIDAstra(pg, idObjeto)
    Set ref = ObjetoPorIDAstra(pg, idReferencia)

    If s Is Nothing Then Exit Sub
    If ref Is Nothing Then Exit Sub

    novoX = ref.RightX + distancia + (s.SizeWidth / 2)

    s.SetPosition novoX, s.CenterY

End Sub

Sub MargemPaginaAstra( _
    ByVal pg As Page, _
    ByVal idObjeto As String, _
    ByVal lado As String, _
    ByVal margem As Double)

    Dim s As Shape
    Dim novoX As Double
    Dim novoY As Double

    Set s = ObjetoPorIDAstra(pg, idObjeto)

    If s Is Nothing Then Exit Sub

    novoX = s.CenterX
    novoY = s.CenterY

    Select Case UCase(Trim(lado))

        Case "LEFT"
            novoX = margem + s.SizeWidth / 2

        Case "RIGHT"
            novoX = pg.SizeWidth - margem - s.SizeWidth / 2

        Case "BOTTOM"
            novoY = margem + s.SizeHeight / 2

        Case "TOP"
            novoY = pg.SizeHeight - margem - s.SizeHeight / 2

    End Select

    s.SetPosition novoX, novoY

End Sub

Sub ManterJuntosAstra( _
    ByVal pg As Page, _
    ByVal listaIDs As String, _
    ByVal novoX As Double, _
    ByVal novoY As Double)

    Dim ids() As String
    Dim i As Long
    Dim s As Shape
    Dim conjunto As New ShapeRange
    
    listaIDs = ResolverSeletorAstra(listaIDs)

    If Trim(listaIDs) = "" Then Exit Sub

    ids = Split(listaIDs, ",")

    For i = LBound(ids) To UBound(ids)

        Set s = ObjetoPorIDAstra(pg, Trim(ids(i)))

        If Not s Is Nothing Then
            conjunto.Add s
        End If

    Next i

    If conjunto.Count = 0 Then Exit Sub

    ActiveDocument.ReferencePoint = cdrCenter

    conjunto.SetPosition novoX, novoY

End Sub

Sub AcimaDeAstra( _
    ByVal pg As Page, _
    ByVal idObjeto As String, _
    ByVal idReferencia As String, _
    ByVal distancia As Double)

    Dim s As Shape
    Dim ref As Shape
    Dim novoY As Double

    Set s = ObjetoPorIDAstra(pg, idObjeto)
    Set ref = ObjetoPorIDAstra(pg, idReferencia)

    If s Is Nothing Then Exit Sub
    If ref Is Nothing Then Exit Sub

    novoY = ref.TopY + distancia + (s.SizeHeight / 2)

    s.SetPosition s.CenterX, novoY

End Sub


Sub DistribuirAstra( _
    ByVal pg As Page, _
    ByVal listaIDs As String, _
    ByVal modo As String)

    Dim ids() As String
    Dim objetos() As Shape

    Dim i As Long
    Dim j As Long
    Dim qtd As Long

    Dim temp As Shape

    Dim minimo As Double
    Dim maximo As Double
    Dim passo As Double
    Dim posicao As Double
    
    listaIDs = ResolverSeletorAstra(listaIDs)

    If Trim(listaIDs) = "" Then Exit Sub

    ids = Split(listaIDs, ",")

    ReDim objetos(LBound(ids) To UBound(ids))

    qtd = 0

    For i = LBound(ids) To UBound(ids)

        Set objetos(i) = ObjetoPorIDAstra(pg, Trim(ids(i)))

        If Not objetos(i) Is Nothing Then
            qtd = qtd + 1
        End If

    Next i

    If qtd < 3 Then Exit Sub

    If UCase(Trim(modo)) = "HORIZONTAL" Then

        ' ordena pela posição X
        For i = LBound(objetos) To UBound(objetos) - 1
            For j = i + 1 To UBound(objetos)

                If Not objetos(i) Is Nothing And Not objetos(j) Is Nothing Then

                    If objetos(i).CenterX > objetos(j).CenterX Then

                        Set temp = objetos(i)
                        Set objetos(i) = objetos(j)
                        Set objetos(j) = temp

                    End If

                End If

            Next j
        Next i

        minimo = objetos(LBound(objetos)).CenterX
        maximo = objetos(UBound(objetos)).CenterX

        passo = (maximo - minimo) / (qtd - 1)

        posicao = minimo

        For i = LBound(objetos) To UBound(objetos)

            If Not objetos(i) Is Nothing Then

                objetos(i).SetPosition posicao, objetos(i).CenterY

                posicao = posicao + passo

            End If

        Next i

    ElseIf UCase(Trim(modo)) = "VERTICAL" Then

        ' ordena pela posição Y
        For i = LBound(objetos) To UBound(objetos) - 1
            For j = i + 1 To UBound(objetos)

                If Not objetos(i) Is Nothing And Not objetos(j) Is Nothing Then

                    If objetos(i).CenterY > objetos(j).CenterY Then

                        Set temp = objetos(i)
                        Set objetos(i) = objetos(j)
                        Set objetos(j) = temp

                    End If

                End If

            Next j
        Next i

        minimo = objetos(LBound(objetos)).CenterY
        maximo = objetos(UBound(objetos)).CenterY

        passo = (maximo - minimo) / (qtd - 1)

        posicao = minimo

        For i = LBound(objetos) To UBound(objetos)

            If Not objetos(i) Is Nothing Then

                objetos(i).SetPosition objetos(i).CenterX, posicao

                posicao = posicao + passo

            End If

        Next i

    End If

End Sub

Function ExpandirSeletorAstra(ByVal seletor As String) As String

    Dim caminho As String
    Dim arq As Integer
    Dim linha As String
    Dim partes() As String
    Dim roleDesejado As String

    seletor = Trim(seletor)

    ' Se não começar com @, já é um ID normal
    If Left(seletor, 1) <> "@" Then
        ExpandirSeletorAstra = seletor
        Exit Function
    End If

    roleDesejado = UCase(Trim(Mid(seletor, 2)))

    caminho = "C:\CorelAI\roles.txt"

    If Dir(caminho) = "" Then
        ExpandirSeletorAstra = ""
        Exit Function
    End If

    arq = FreeFile
    Open caminho For Input As #arq

    Do Until EOF(arq)

        Line Input #arq, linha
        linha = Trim(linha)

        If linha <> "" Then

            partes = Split(linha, "|")

            If UBound(partes) >= 2 Then

                If UCase(Trim(partes(0))) = "ROLE" Then

                    If UCase(Trim(partes(1))) = roleDesejado Then

                        ExpandirSeletorAstra = Trim(partes(2))
                        Close #arq
                        Exit Function

                    End If

                End If

            End If

        End If

    Loop

    Close #arq

    ExpandirSeletorAstra = ""

End Function
Function PrimeiroIDAstra(ByVal seletor As String) As String

    Dim lista As String
    Dim partes() As String

    lista = ResolverSeletorAstra(seletor)

    If Trim(lista) = "" Then
        PrimeiroIDAstra = ""
        Exit Function
    End If

    partes = Split(lista, ",")

    PrimeiroIDAstra = Trim(partes(0))

End Function

Sub DefinirZonaAstra( _
    ByVal nomeZona As String, _
    ByVal xPct As Double, _
    ByVal yPct As Double, _
    ByVal larguraPct As Double, _
    ByVal alturaPct As Double)

    Dim dados As String

    If ZonasAstra Is Nothing Then
        Set ZonasAstra = CreateObject("Scripting.Dictionary")
    End If

    nomeZona = UCase(Trim(nomeZona))

    dados = _
        CStr(xPct) & "|" & _
        CStr(yPct) & "|" & _
        CStr(larguraPct) & "|" & _
        CStr(alturaPct)

    ZonasAstra(nomeZona) = dados

End Sub

Function DadosZonaAstra( _
    ByVal pg As Page, _
    ByVal nomeZona As String, _
    ByRef x As Double, _
    ByRef y As Double, _
    ByRef largura As Double, _
    ByRef altura As Double) As Boolean

    Dim dados() As String
    Dim bruto As String

    DadosZonaAstra = False

    If ZonasAstra Is Nothing Then Exit Function

    nomeZona = UCase(Trim(nomeZona))

    If Not ZonasAstra.Exists(nomeZona) Then Exit Function

    bruto = ZonasAstra(nomeZona)
    dados = Split(bruto, "|")

    If UBound(dados) < 3 Then Exit Function

    x = pg.SizeWidth * NumeroAstra(dados(0)) / 100
    y = pg.SizeHeight * NumeroAstra(dados(1)) / 100

    largura = pg.SizeWidth * NumeroAstra(dados(2)) / 100
    altura = pg.SizeHeight * NumeroAstra(dados(3)) / 100

    DadosZonaAstra = True

End Function

Sub PosicionarNaZonaAstra( _
    ByVal pg As Page, _
    ByVal seletor As String, _
    ByVal nomeZona As String, _
    ByVal alinhamentoH As String, _
    ByVal alinhamentoV As String, _
    ByVal fatorMaximo As Double)

    Dim lista As String
    Dim ids() As String

    Dim conjunto As New ShapeRange
    Dim s As Shape

    Dim i As Long

    Dim zonaX As Double
    Dim zonaY As Double
    Dim zonaLargura As Double
    Dim zonaAltura As Double

    Dim larguraMax As Double
    Dim alturaMax As Double

    Dim escalaL As Double
    Dim escalaA As Double
    Dim escala As Double

    Dim novoX As Double
    Dim novoY As Double

    listaIDs = ResolverSeletorAstra(listaIDs)

    If Trim(lista) = "" Then Exit Sub

    ids = Split(lista, ",")

    For i = LBound(ids) To UBound(ids)

        Set s = ObjetoPorIDAstra(pg, Trim(ids(i)))

        If Not s Is Nothing Then
            conjunto.Add s
        End If

    Next i

    If conjunto.Count = 0 Then Exit Sub

    If Not DadosZonaAstra( _
        pg, _
        nomeZona, _
        zonaX, _
        zonaY, _
        zonaLargura, _
        zonaAltura) Then Exit Sub

    If fatorMaximo <= 0 Then fatorMaximo = 1

    larguraMax = zonaLargura * fatorMaximo
    alturaMax = zonaAltura * fatorMaximo

    escalaL = larguraMax / conjunto.SizeWidth
    escalaA = alturaMax / conjunto.SizeHeight

    If escalaL < escalaA Then
        escala = escalaL
    Else
        escala = escalaA
    End If

    ' Não aumenta automaticamente.
    ' Só reduz se estiver maior que a zona.
    If escala < 1 Then

        ActiveDocument.ReferencePoint = cdrCenter

        conjunto.SetSize _
            conjunto.SizeWidth * escala, _
            conjunto.SizeHeight * escala

    End If

    ' -----------------------
    ' ALINHAMENTO HORIZONTAL
    ' -----------------------

    Select Case UCase(Trim(alinhamentoH))

        Case "LEFT"

            novoX = _
                zonaX + _
                (conjunto.SizeWidth / 2)

        Case "RIGHT"

            novoX = _
                zonaX + _
                zonaLargura - _
                (conjunto.SizeWidth / 2)

        Case Else

            novoX = _
                zonaX + _
                (zonaLargura / 2)

    End Select

    ' -----------------------
    ' ALINHAMENTO VERTICAL
    ' -----------------------

    Select Case UCase(Trim(alinhamentoV))

        Case "BOTTOM"

            novoY = _
                zonaY + _
                (conjunto.SizeHeight / 2)

        Case "TOP"

            novoY = _
                zonaY + _
                zonaAltura - _
                (conjunto.SizeHeight / 2)

        Case Else

            novoY = _
                zonaY + _
                (zonaAltura / 2)

    End Select

    ActiveDocument.ReferencePoint = cdrCenter

    conjunto.SetPosition novoX, novoY

End Sub

Function ValidarLayoutAstra(ByVal pg As Page, ByVal caminho As String) As Boolean

    Dim arq As Integer
    Dim linha As String
    Dim partes() As String
    Dim erros As String
    Dim avisos As String
    Dim seletor As String
    Dim id As String
    Dim lista As String
    Dim ids() As String
    Dim i As Long
    Dim s As Shape
    Dim nomeZona As String

    ValidarLayoutAstra = False

    If Dir(caminho) = "" Then
        MsgBox "Arquivo de layout não encontrado.", vbExclamation, "Corel AI"
        Exit Function
    End If

    If ZonasAstra Is Nothing Then
        Set ZonasAstra = CreateObject("Scripting.Dictionary")
    Else
        ZonasAstra.RemoveAll
    End If

    arq = FreeFile
    Open caminho For Input As #arq

    Do Until EOF(arq)

        Line Input #arq, linha
        linha = Trim(linha)

        If linha <> "" Then

            partes = Split(linha, "|")

            Select Case UCase(Trim(partes(0)))
            
            Case "NOCOLLIDE"

    If UBound(partes) < 4 Then

        erros = erros & _
            "NOCOLLIDE incompleto: " & linha & vbCrLf

    Else

        lista = ResolverSeletorAstra(partes(1))

        If Trim(lista) = "" Then
            erros = erros & _
                "Seletor inexistente em NOCOLLIDE: " & _
                partes(1) & vbCrLf
        End If

        lista = ResolverSeletorAstra(partes(2))

        If Trim(lista) = "" Then
            erros = erros & _
                "Seletor inexistente em NOCOLLIDE: " & _
                partes(2) & vbCrLf
        End If

    End If
            
            Case "SAFEAREA"

    If UBound(partes) < 2 Then
        erros = erros & _
            "SAFEAREA incompleto: " & linha & vbCrLf
    End If


Case "AUTOFIT"

    If UBound(partes) < 1 Then

        erros = erros & _
            "AUTOFIT incompleto: " & linha & vbCrLf

    Else

        lista = ResolverSeletorAstra(partes(1))

        If Trim(lista) = "" Then
            erros = erros & _
                "Seletor inexistente em AUTOFIT: " & _
                partes(1) & vbCrLf
        End If

    End If
            
            Case "GAP"

            If UBound(partes) < 5 Then

            erros = erros & _
            "GAP incompleto: " & linha & vbCrLf

            Else

            lista = ExpandirSeletorAstra(partes(1))

            If Trim(lista) = "" Then
            erros = erros & _
                "Referência inexistente em GAP: " & _
                partes(1) & vbCrLf
            End If

            lista = ExpandirSeletorAstra(partes(2))

            If Trim(lista) = "" Then
            erros = erros & _
                "Objeto móvel inexistente em GAP: " & _
                partes(2) & vbCrLf
            End If

            Select Case UCase(Trim(partes(3)))

            Case "LEFT", "RIGHT", "ABOVE", "BELOW"

            Case Else

                erros = erros & _
                    "Direção inválida em GAP: " & _
                    partes(3) & vbCrLf

            End Select

            End If

                Case "PAGE"

                    If UBound(partes) < 2 Then
                        erros = erros & "PAGE incompleto: " & linha & vbCrLf
                    End If


                Case "ZONE"

                    If UBound(partes) < 5 Then

                        erros = erros & "ZONE incompleta: " & linha & vbCrLf

                    Else

                        DefinirZonaAstra _
                            partes(1), _
                            NumeroAstra(partes(2)), _
                            NumeroAstra(partes(3)), _
                            NumeroAstra(partes(4)), _
                            NumeroAstra(partes(5))

                    End If


                Case "FILLPAGE", "OBJECT", "STRETCH", "SCALE", "MARGIN", "ABOVE", "BELOW", "RIGHTOF"

    If UBound(partes) < 1 Then

        erros = erros & _
            "Comando incompleto: " & linha & vbCrLf

    Else

        seletor = Trim(partes(1))
        lista = ResolverSeletorAstra(seletor)

        If Trim(lista) = "" Then

            erros = erros & _
                "Seletor inexistente: " & seletor & vbCrLf

        Else

            ids = Split(lista, ",")

            For i = LBound(ids) To UBound(ids)

                Set s = ObjetoPorIDAstra(pg, Trim(ids(i)))

                If s Is Nothing Then

                    erros = erros & _
                        "ID inexistente em " & seletor & _
                        ": " & Trim(ids(i)) & vbCrLf

                End If

            Next i

        End If

    End If


                Case "SET", "KEEPTOGETHER", "DISTRIBUTE", "ALIGN"

                    If UBound(partes) < 1 Then

                        erros = erros & "Comando incompleto: " & linha & vbCrLf

                    Else

                        lista = ExpandirSeletorAstra(partes(1))

                        If Trim(lista) = "" Then

                            erros = erros & "Seletor vazio ou role inexistente: " & partes(1) & vbCrLf

                        Else

                            ids = Split(lista, ",")

                            For i = LBound(ids) To UBound(ids)

                                Set s = ObjetoPorIDAstra(pg, Trim(ids(i)))

                                If s Is Nothing Then
                                    erros = erros & "ID inexistente: " & Trim(ids(i)) & vbCrLf
                                End If

                            Next i

                        End If

                    End If


                Case "FITINSIDE"

                    If UBound(partes) < 3 Then

                        erros = erros & "FITINSIDE incompleto: " & linha & vbCrLf

                    Else

                        lista = ExpandirSeletorAstra(partes(1))

                        If Trim(lista) = "" Then
                            erros = erros & "Conteúdo inexistente em FITINSIDE: " & partes(1) & vbCrLf
                        End If

                        id = PrimeiroIDAstra(partes(2))

                        If Trim(id) = "" Then
                            erros = erros & "Container inexistente: " & partes(2) & vbCrLf
                        Else

                            Set s = ObjetoPorIDAstra(pg, id)

                            If s Is Nothing Then
                                erros = erros & "Container inválido: " & partes(2) & vbCrLf
                            End If

                        End If

                    End If


                Case "PLACEZONE"

                    If UBound(partes) < 5 Then

                        erros = erros & "PLACEZONE incompleto: " & linha & vbCrLf

                    Else

                        nomeZona = UCase(Trim(partes(2)))

                        If Not ZonasAstra.Exists(nomeZona) Then
                            erros = erros & "Zona inexistente: " & nomeZona & vbCrLf
                        End If

                        lista = ExpandirSeletorAstra(partes(1))

                        If Trim(lista) = "" Then
                            erros = erros & "Seletor inexistente em PLACEZONE: " & partes(1) & vbCrLf
                        End If

                    End If


                Case Else

                    avisos = avisos & "Comando desconhecido: " & linha & vbCrLf

            End Select

        End If

    Loop

    Close #arq

    If erros <> "" Then

        MsgBox _
            "LAYOUT NÃO APLICADO." & vbCrLf & vbCrLf & _
            "Erros encontrados:" & vbCrLf & vbCrLf & _
            erros & _
            IIf(avisos <> "", vbCrLf & "Avisos:" & vbCrLf & avisos, ""), _
            vbCritical, _
            "Validação Corel AI"

        Exit Function

    End If

    If avisos <> "" Then

        MsgBox _
            "Validação concluída com avisos:" & vbCrLf & vbCrLf & _
            avisos, _
            vbExclamation, _
            "Corel AI"

    End If

    ValidarLayoutAstra = True

End Function

Sub ValidarResultadoAstra(ByVal pg As Page)

    Dim sr As ShapeRange
    Dim s As Shape
    Dim i As Long
    Dim avisos As String
    Dim tolerancia As Double

    tolerancia = 0.5

    Set sr = pg.Shapes.All

    For i = 1 To sr.Count

        Set s = sr(i)

        If s.LeftX < -tolerancia Then
            avisos = avisos & _
                "Objeto " & i & " ultrapassa a borda esquerda." & vbCrLf
        End If

        If s.RightX > pg.SizeWidth + tolerancia Then
            avisos = avisos & _
                "Objeto " & i & " ultrapassa a borda direita." & vbCrLf
        End If

        If s.BottomY < -tolerancia Then
            avisos = avisos & _
                "Objeto " & i & " ultrapassa a borda inferior." & vbCrLf
        End If

        If s.TopY > pg.SizeHeight + tolerancia Then
            avisos = avisos & _
                "Objeto " & i & " ultrapassa a borda superior." & vbCrLf
        End If

    Next i

    If avisos <> "" Then

        MsgBox _
            "Layout aplicado, mas há elementos fora da página:" & vbCrLf & vbCrLf & _
            avisos, _
            vbExclamation, _
            "Corel AI"

    Else

        MsgBox _
            "Layout aplicado e validado." & vbCrLf & _
            "Nenhum objeto principal ultrapassa a página.", _
            vbInformation, _
            "Corel AI"

    End If

End Sub

Function AreaPercentualAstra( _
    ByVal pg As Page, _
    ByVal seletor As String) As Double

    Dim lista As String
    Dim ids() As String
    Dim conjunto As New ShapeRange
    Dim s As Shape
    Dim i As Long

    lista = ExpandirSeletorAstra(seletor)

    If Trim(lista) = "" Then
        AreaPercentualAstra = 0
        Exit Function
    End If

    ids = Split(lista, ",")

    For i = LBound(ids) To UBound(ids)

        Set s = ObjetoPorIDAstra(pg, Trim(ids(i)))

        If Not s Is Nothing Then
            conjunto.Add s
        End If

    Next i

    If conjunto.Count = 0 Then
        AreaPercentualAstra = 0
        Exit Function
    End If

    AreaPercentualAstra = _
        ((conjunto.SizeWidth * conjunto.SizeHeight) / _
        (pg.SizeWidth * pg.SizeHeight)) * 100

End Function

Function RangeDoSeletorAstra( _
    ByVal pg As Page, _
    ByVal seletor As String) As ShapeRange

    Dim lista As String
    Dim ids() As String
    Dim conjunto As New ShapeRange
    Dim s As Shape
    Dim i As Long

    lista = ResolverSeletorAstra(seletor)

    If Trim(lista) = "" Then
        Set RangeDoSeletorAstra = conjunto
        Exit Function
    End If

    ids = Split(lista, ",")

    For i = LBound(ids) To UBound(ids)

        Set s = ObjetoPorIDAstra(pg, Trim(ids(i)))

        If Not s Is Nothing Then
            conjunto.Add s
        End If

    Next i

    Set RangeDoSeletorAstra = conjunto

End Function

Sub AvaliarLayoutSemanticoAstra(ByVal pg As Page)

    Dim relatorio As String
    Dim avisos As String

    Dim areaProduto As Double
    Dim areaPreco As Double
    Dim areaTitulo As Double
    Dim areaCampanha As Double

    Dim produto As ShapeRange
    Dim preco As ShapeRange
    Dim titulo As ShapeRange
    Dim fundo As ShapeRange
    Dim faixa As ShapeRange

    relatorio = "ANÁLISE SEMÂNTICA" & vbCrLf & vbCrLf

    areaProduto = AreaPercentualAstra(pg, "@PRODUCT")
    areaPreco = AreaPercentualAstra(pg, "@PRICE_PRIMARY")
    areaTitulo = AreaPercentualAstra(pg, "@TITLE")
    areaCampanha = AreaPercentualAstra(pg, "@CAMPAIGN")

    relatorio = relatorio & _
        "Produto: " & Format(areaProduto, "0.0") & "% da página" & vbCrLf & _
        "Preço principal: " & Format(areaPreco, "0.0") & "%" & vbCrLf & _
        "Título: " & Format(areaTitulo, "0.0") & "%" & vbCrLf & _
        "Campanha: " & Format(areaCampanha, "0.0") & "%" & vbCrLf & vbCrLf

    ' -------------------------
    ' PRODUTO
    ' -------------------------

    If areaProduto > 0 Then

        If areaProduto < 8 Then
            avisos = avisos & _
                "Produto pode estar pequeno demais." & vbCrLf
        End If

        If areaProduto > 45 Then
            avisos = avisos & _
                "Produto pode estar dominando demais a composição." & vbCrLf
        End If

    End If

    ' -------------------------
    ' PREÇO PRINCIPAL
    ' -------------------------

    If areaPreco > 0 Then

        If areaPreco < 5 Then
            avisos = avisos & _
                "Preço principal pode estar pequeno demais." & vbCrLf
        End If

        If areaProduto > 0 Then

            If areaPreco < areaProduto * 0.25 Then
                avisos = avisos & _
                    "Preço principal está muito fraco em relação ao produto." & vbCrLf
            End If

        End If

    End If

    ' -------------------------
    ' TÍTULO
    ' -------------------------

    If areaTitulo > 0 And areaTitulo < 0.5 Then
        avisos = avisos & _
            "Título pode estar pequeno demais." & vbCrLf
    End If

    ' -------------------------
    ' FUNDO
    ' -------------------------

    Set fundo = RangeDoSeletorAstra(pg, "@BACKGROUND")

    If fundo.Count > 0 Then

        If fundo.LeftX > 1 Or _
           fundo.RightX < pg.SizeWidth - 1 Or _
           fundo.BottomY > 1 Or _
           fundo.TopY < pg.SizeHeight - 1 Then

            avisos = avisos & _
                "O fundo não cobre completamente a página." & vbCrLf

        End If

    End If

    ' -------------------------
    ' TÍTULO x FAIXA
    ' -------------------------

    Set titulo = RangeDoSeletorAstra(pg, "@TITLE")
    Set faixa = RangeDoSeletorAstra(pg, "@BAND_PRIMARY")

    If titulo.Count > 0 And faixa.Count > 0 Then

        If titulo.CenterY < faixa.CenterY Then
            avisos = avisos & _
                "O título está abaixo do centro da faixa principal." & vbCrLf
        End If

    End If

    ' -------------------------
    ' RESULTADO
    ' -------------------------

    If avisos = "" Then

        relatorio = relatorio & _
            "Nenhum problema semântico básico detectado."

    Else

        relatorio = relatorio & _
            "Pontos para revisar:" & vbCrLf & vbCrLf & avisos

    End If

    MsgBox relatorio, vbInformation, "Corel AI"

End Sub

Sub SalvarAnaliseSemanticaAstra(ByVal pg As Page)

    Dim arq As Integer
    Dim caminho As String

    caminho = "C:\CorelAI\analise_layout.txt"

    arq = FreeFile
    Open caminho For Output As #arq

    Print #arq, "ANALISE_LAYOUT"
    Print #arq, ""
    Print #arq, "PRODUCT_AREA_PCT=" & _
        NumeroTexto(AreaPercentualAstra(pg, "@PRODUCT"))

    Print #arq, "PRICE_PRIMARY_AREA_PCT=" & _
        NumeroTexto(AreaPercentualAstra(pg, "@PRICE_PRIMARY"))

    Print #arq, "TITLE_AREA_PCT=" & _
        NumeroTexto(AreaPercentualAstra(pg, "@TITLE"))

    Print #arq, "CAMPAIGN_AREA_PCT=" & _
        NumeroTexto(AreaPercentualAstra(pg, "@CAMPAIGN"))

    Close #arq

End Sub

Function SobreposicaoAstra( _
    ByVal a As ShapeRange, _
    ByVal b As ShapeRange) As Double

    Dim esquerda As Double
    Dim direita As Double
    Dim topo As Double
    Dim base As Double

    Dim largura As Double
    Dim altura As Double

    Dim areaIntersecao As Double
    Dim areaMenor As Double

    SobreposicaoAstra = 0

    If a.Count = 0 Or b.Count = 0 Then Exit Function

    esquerda = a.LeftX
    If b.LeftX > esquerda Then esquerda = b.LeftX

    direita = a.RightX
    If b.RightX < direita Then direita = b.RightX

    base = a.BottomY
    If b.BottomY > base Then base = b.BottomY

    topo = a.TopY
    If b.TopY < topo Then topo = b.TopY

    largura = direita - esquerda
    altura = topo - base

    If largura <= 0 Or altura <= 0 Then Exit Function

    areaIntersecao = largura * altura

    areaMenor = a.SizeWidth * a.SizeHeight

    If (b.SizeWidth * b.SizeHeight) < areaMenor Then
        areaMenor = b.SizeWidth * b.SizeHeight
    End If

    If areaMenor <= 0 Then Exit Function

    SobreposicaoAstra = (areaIntersecao / areaMenor) * 100

End Function

Sub AvaliarColisoesAstra(ByVal pg As Page)

    Dim produto As ShapeRange
    Dim preco As ShapeRange
    Dim titulo As ShapeRange
    Dim campanha As ShapeRange
    Dim faixa As ShapeRange

    Dim relatorio As String
    Dim pct As Double

    Set produto = RangeDoSeletorAstra(pg, "@PRODUCT")
    Set preco = RangeDoSeletorAstra(pg, "@PRICE_PRIMARY")
    Set titulo = RangeDoSeletorAstra(pg, "@TITLE")
    Set campanha = RangeDoSeletorAstra(pg, "@CAMPAIGN")
    Set faixa = RangeDoSeletorAstra(pg, "@BAND_PRIMARY")

    relatorio = ""

    pct = SobreposicaoAstra(produto, preco)

    If pct > 10 Then
        relatorio = relatorio & _
            "PRODUCT x PRICE_PRIMARY = " & _
            Format(pct, "0.0") & "%" & vbCrLf
    End If

    pct = SobreposicaoAstra(produto, titulo)

    If pct > 10 Then
        relatorio = relatorio & _
            "PRODUCT x TITLE = " & _
            Format(pct, "0.0") & "%" & vbCrLf
    End If

    pct = SobreposicaoAstra(produto, campanha)

    If pct > 15 Then
        relatorio = relatorio & _
            "PRODUCT x CAMPAIGN = " & _
            Format(pct, "0.0") & "%" & vbCrLf
    End If

    pct = SobreposicaoAstra(preco, titulo)

    If pct > 10 Then
        relatorio = relatorio & _
            "PRICE_PRIMARY x TITLE = " & _
            Format(pct, "0.0") & "%" & vbCrLf
    End If

    If relatorio <> "" Then

        MsgBox _
            "Possíveis colisões visuais:" & vbCrLf & vbCrLf & _
            relatorio, _
            vbExclamation, _
            "Corel AI"

    End If

End Sub

Sub SalvarRelatorioLayoutAstra(ByVal pg As Page)

    Dim caminho As String
    Dim arq As Integer

    Dim produto As ShapeRange
    Dim preco As ShapeRange
    Dim titulo As ShapeRange
    Dim campanha As ShapeRange

    caminho = "C:\CorelAI\relatorio_layout.txt"

    Set produto = RangeDoSeletorAstra(pg, "@PRODUCT")
    Set preco = RangeDoSeletorAstra(pg, "@PRICE_PRIMARY")
    Set titulo = RangeDoSeletorAstra(pg, "@TITLE")
    Set campanha = RangeDoSeletorAstra(pg, "@CAMPAIGN")

    arq = FreeFile
    Open caminho For Output As #arq

    Print #arq, "RELATORIO_LAYOUT"
    Print #arq, ""

    Print #arq, "PRODUCT_AREA_PCT=" & _
        NumeroTexto(AreaPercentualAstra(pg, "@PRODUCT"))

    Print #arq, "PRICE_PRIMARY_AREA_PCT=" & _
        NumeroTexto(AreaPercentualAstra(pg, "@PRICE_PRIMARY"))

    Print #arq, "TITLE_AREA_PCT=" & _
        NumeroTexto(AreaPercentualAstra(pg, "@TITLE"))

    Print #arq, ""

    Print #arq, "OVERLAP_PRODUCT_PRICE=" & _
        NumeroTexto(SobreposicaoAstra(produto, preco))

    Print #arq, "OVERLAP_PRODUCT_TITLE=" & _
        NumeroTexto(SobreposicaoAstra(produto, titulo))

    Print #arq, "OVERLAP_PRODUCT_CAMPAIGN=" & _
        NumeroTexto(SobreposicaoAstra(produto, campanha))

    Print #arq, "OVERLAP_PRICE_TITLE=" & _
        NumeroTexto(SobreposicaoAstra(preco, titulo))

    Close #arq

End Sub

Sub AplicarTesteAstra()

    AplicarLayoutArquivo "C:\CorelAI\layout_teste_07.txt"

End Sub

Sub GapAstra( _
    ByVal pg As Page, _
    ByVal seletorReferencia As String, _
    ByVal seletorMover As String, _
    ByVal direcao As String, _
    ByVal distancia As Double, _
    ByVal unidade As String)

    Dim referencia As ShapeRange
    Dim mover As ShapeRange

    Dim gap As Double
    Dim novoX As Double
    Dim novoY As Double

    Set referencia = RangeDoSeletorAstra(pg, seletorReferencia)
    Set mover = RangeDoSeletorAstra(pg, seletorMover)

    If referencia.Count = 0 Then Exit Sub
    If mover.Count = 0 Then Exit Sub

    unidade = UCase(Trim(unidade))

    Select Case unidade

        Case "PCTW"
            gap = pg.SizeWidth * distancia / 100

        Case "PCTH"
            gap = pg.SizeHeight * distancia / 100

        Case Else
            gap = distancia

    End Select

    direcao = UCase(Trim(direcao))

    novoX = mover.CenterX
    novoY = mover.CenterY

    Select Case direcao

        Case "RIGHT"

            novoX = _
                referencia.RightX + _
                gap + _
                mover.SizeWidth / 2

        Case "LEFT"

            novoX = _
                referencia.LeftX - _
                gap - _
                mover.SizeWidth / 2

        Case "ABOVE"

            novoY = _
                referencia.TopY + _
                gap + _
                mover.SizeHeight / 2

        Case "BELOW"

            novoY = _
                referencia.BottomY - _
                gap - _
                mover.SizeHeight / 2

        Case Else

            Exit Sub

    End Select

    ActiveDocument.ReferencePoint = cdrCenter
    mover.SetPosition novoX, novoY

End Sub

Function ExpandirBlocoAstra(ByVal seletor As String) As String

    Dim caminho As String
    Dim arq As Integer
    Dim linha As String
    Dim partes() As String

    Dim nomeDesejado As String
    Dim conteudo As String
    Dim itens() As String
    Dim resultado As String
    Dim expandido As String
    Dim i As Long

    seletor = Trim(seletor)

    If Left(seletor, 1) <> "#" Then
        ExpandirBlocoAstra = seletor
        Exit Function
    End If

    nomeDesejado = UCase(Trim(Mid(seletor, 2)))
    caminho = "C:\CorelAI\blocks.txt"

    If Dir(caminho) = "" Then
        ExpandirBlocoAstra = ""
        Exit Function
    End If

    arq = FreeFile
    Open caminho For Input As #arq

    Do Until EOF(arq)

        Line Input #arq, linha
        linha = Trim(linha)

        If linha <> "" Then

            partes = Split(linha, "|")

            If UBound(partes) >= 2 Then

                If UCase(Trim(partes(0))) = "BLOCK" Then

                    If UCase(Trim(partes(1))) = nomeDesejado Then

                        conteudo = Trim(partes(2))
                        Close #arq

                        itens = Split(conteudo, ",")

                        For i = LBound(itens) To UBound(itens)

                            expandido = ExpandirSeletorAstra(Trim(itens(i)))

                            If Trim(expandido) <> "" Then

                                If resultado <> "" Then
                                    resultado = resultado & ","
                                End If

                                resultado = resultado & expandido

                            End If

                        Next i

                        ExpandirBlocoAstra = resultado
                        Exit Function

                    End If

                End If

            End If

        End If

    Loop

    Close #arq

    ExpandirBlocoAstra = ""

End Function

Function ResolverSeletorAstra(ByVal seletor As String) As String

    seletor = Trim(seletor)

    If seletor = "" Then
        ResolverSeletorAstra = ""
        Exit Function
    End If

    Select Case Left(seletor, 1)

        Case "@"
            ResolverSeletorAstra = ExpandirSeletorAstra(seletor)

        Case "#"
            ResolverSeletorAstra = ExpandirBlocoAstra(seletor)

        Case Else
            ResolverSeletorAstra = seletor

    End Select
    
End Function

Sub MargemSeletorAstra( _
    ByVal pg As Page, _
    ByVal seletor As String, _
    ByVal lado As String, _
    ByVal margem As Double)

    Dim conjunto As ShapeRange
    Dim novoX As Double
    Dim novoY As Double

    Set conjunto = RangeDoSeletorAstra(pg, seletor)

    If conjunto.Count = 0 Then Exit Sub

    novoX = conjunto.CenterX
    novoY = conjunto.CenterY

    Select Case UCase(Trim(lado))

        Case "LEFT"
            novoX = margem + conjunto.SizeWidth / 2

        Case "RIGHT"
            novoX = pg.SizeWidth - margem - conjunto.SizeWidth / 2

        Case "BOTTOM"
            novoY = margem + conjunto.SizeHeight / 2

        Case "TOP"
            novoY = pg.SizeHeight - margem - conjunto.SizeHeight / 2

        Case Else
            Exit Sub

    End Select

    ActiveDocument.ReferencePoint = cdrCenter
    conjunto.SetPosition novoX, novoY

End Sub

Sub DefinirSafeAreaAstra( _
    ByVal pg As Page, _
    ByVal valor As Double, _
    ByVal unidade As String)

    Select Case UCase(Trim(unidade))

        Case "PCT"
            MargemSeguraAstra = _
                (pg.SizeWidth * valor / 100)

        Case "PCTH"
            MargemSeguraAstra = _
                (pg.SizeHeight * valor / 100)

        Case Else
            MargemSeguraAstra = valor

    End Select

End Sub

Sub AutoFitAstra( _
    ByVal pg As Page, _
    ByVal seletor As String)

    Dim conjunto As ShapeRange

    Dim limiteEsquerdo As Double
    Dim limiteDireito As Double
    Dim limiteInferior As Double
    Dim limiteSuperior As Double

    Dim larguraDisponivel As Double
    Dim alturaDisponivel As Double

    Dim escalaL As Double
    Dim escalaA As Double
    Dim escala As Double

    Dim deslocX As Double
    Dim deslocY As Double

    Set conjunto = RangeDoSeletorAstra(pg, seletor)

    If conjunto.Count = 0 Then Exit Sub

    limiteEsquerdo = MargemSeguraAstra
    limiteDireito = pg.SizeWidth - MargemSeguraAstra
    limiteInferior = MargemSeguraAstra
    limiteSuperior = pg.SizeHeight - MargemSeguraAstra

    larguraDisponivel = limiteDireito - limiteEsquerdo
    alturaDisponivel = limiteSuperior - limiteInferior

    ' Primeiro reduz somente se o conjunto for maior
    ' que a área segura.

    If conjunto.SizeWidth > larguraDisponivel Or _
       conjunto.SizeHeight > alturaDisponivel Then

        escalaL = larguraDisponivel / conjunto.SizeWidth
        escalaA = alturaDisponivel / conjunto.SizeHeight

        If escalaL < escalaA Then
            escala = escalaL
        Else
            escala = escalaA
        End If

        ActiveDocument.ReferencePoint = cdrCenter

        conjunto.SetSize _
            conjunto.SizeWidth * escala, _
            conjunto.SizeHeight * escala

    End If

    ' Depois corrige posição sem alterar tamanho.

    deslocX = 0
    deslocY = 0

    If conjunto.LeftX < limiteEsquerdo Then
        deslocX = limiteEsquerdo - conjunto.LeftX
    End If

    If conjunto.RightX + deslocX > limiteDireito Then
        deslocX = deslocX - _
            ((conjunto.RightX + deslocX) - limiteDireito)
    End If

    If conjunto.BottomY < limiteInferior Then
        deslocY = limiteInferior - conjunto.BottomY
    End If

    If conjunto.TopY + deslocY > limiteSuperior Then
        deslocY = deslocY - _
            ((conjunto.TopY + deslocY) - limiteSuperior)
    End If

    ActiveDocument.ReferencePoint = cdrCenter

    conjunto.SetPosition _
        conjunto.CenterX + deslocX, _
        conjunto.CenterY + deslocY

End Sub

Sub NoCollideAstra( _
    ByVal pg As Page, _
    ByVal seletorA As String, _
    ByVal seletorB As String, _
    ByVal direcao As String, _
    ByVal gap As Double)

    Dim a As ShapeRange
    Dim b As ShapeRange
    Dim moverX As Double
    Dim moverY As Double

    Set a = RangeDoSeletorAstra(pg, seletorA)
    Set b = RangeDoSeletorAstra(pg, seletorB)

    If a.Count = 0 Or b.Count = 0 Then Exit Sub

    ' Se não existe interseção, não faz nada.
    If a.RightX + gap <= b.LeftX Or _
       b.RightX + gap <= a.LeftX Or _
       a.TopY + gap <= b.BottomY Or _
       b.TopY + gap <= a.BottomY Then
        Exit Sub
    End If

    moverX = b.CenterX
    moverY = b.CenterY

    Select Case UCase(Trim(direcao))

        Case "RIGHT"

            moverX = _
                a.RightX + gap + b.SizeWidth / 2

        Case "LEFT"

            moverX = _
                a.LeftX - gap - b.SizeWidth / 2

        Case "ABOVE"

            moverY = _
                a.TopY + gap + b.SizeHeight / 2

        Case "BELOW"

            moverY = _
                a.BottomY - gap - b.SizeHeight / 2

        Case Else
            Exit Sub

    End Select

    ActiveDocument.ReferencePoint = cdrCenter
    b.SetPosition moverX, moverY

End Sub

Sub AplicarLayoutProducao()

    AplicarLayoutArquivo "C:\CorelAI\layout_producao.txt"

End Sub

Sub PrepararLoteParaAstra()

    Dim sr As ShapeRange
    Dim arq As Integer
    Dim caminho As String

    Dim i As Long
    Dim s As Shape

    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Selecione os grupos das artes do lote.", vbExclamation, "Corel AI"
        Exit Sub
    End If

    Set sr = ActiveSelectionRange
    
    ' Exporta o preview enquanto a seleção original ainda está ativa
    ExportarPreviewLoteParaAstra

    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter

    caminho = "C:\CorelAI\cena_lote.txt"

    arq = FreeFile
    Open caminho For Output As #arq

    Print #arq, "LOTE_COREL_AI"
    Print #arq, "TOTAL_ARTES=" & sr.Count
    Print #arq, ""

    For i = 1 To sr.Count

        Set s = sr(i)

        Print #arq, "=============================="
        Print #arq, "ARTE=" & i
        Print #arq, "ID_TOPO=" & i
        Print #arq, "TIPO=" & TipoDoObjeto(s)

        Print #arq, "X_MM=" & NumeroTexto(s.CenterX)
        Print #arq, "Y_MM=" & NumeroTexto(s.CenterY)
        Print #arq, "LARGURA_MM=" & NumeroTexto(s.SizeWidth)
        Print #arq, "ALTURA_MM=" & NumeroTexto(s.SizeHeight)

        If Trim(s.Name) <> "" Then
            Print #arq, "NOME=" & LimparTextoCena(s.Name)
        End If

        Print #arq, ""

        ExportarInfoGrupoDetalhado _
            s, _
            arq, _
            1, _
            CStr(i)

        Print #arq, ""

    Next i

    Close #arq

    
    
    EnviarPedidoLoteAoAstra

    MsgBox _
        "Lote preparado." & vbCrLf & vbCrLf & _
        "Artes: " & sr.Count & vbCrLf & _
        "Cena: C:\CorelAI\cena_lote.txt" & vbCrLf & _
        "Preview: C:\CorelAI\preview_lote.png", _
        vbInformation, _
        "Corel AI"
        

End Sub
Sub ExportarPreviewLoteParaAstra()

    Dim caminho As String
    Dim larguraPx As Long
    Dim alturaPx As Long
    Dim proporcao As Double
    Dim sr As ShapeRange
    Dim filtro As ExportFilter

    caminho = "C:\CorelAI\preview_lote.png"

    If ActiveSelectionRange.Count = 0 Then
        MsgBox _
            "Selecione as artes do lote antes de exportar o preview.", _
            vbExclamation, _
            "Corel AI"
        Exit Sub
    End If

    Set sr = ActiveSelectionRange

    larguraPx = 2000

    If sr.SizeWidth <= 0 Or sr.SizeHeight <= 0 Then
        MsgBox "A seleção não possui tamanho válido.", vbExclamation, "Corel AI"
        Exit Sub
    End If

    proporcao = sr.SizeHeight / sr.SizeWidth
    alturaPx = CLng(larguraPx * proporcao)

    If alturaPx < 200 Then alturaPx = 200
    If alturaPx > 5000 Then alturaPx = 5000

    ' Garante que a seleção está realmente ativa
    sr.CreateSelection

    Set filtro = ActiveDocument.ExportBitmap( _
        caminho, _
        cdrPNG, _
        cdrSelection, _
        cdrRGBColorImage, _
        larguraPx, _
        alturaPx, _
        96, _
        96, _
        cdrNormalAntiAliasing, _
        False, _
        False, _
        True, _
        False, _
        cdrCompressionNone)

    ' IMPORTANTE: finaliza a gravação do arquivo
    filtro.Finish

    If Dir(caminho) = "" Then
        MsgBox _
            "Falha ao criar:" & vbCrLf & caminho, _
            vbCritical, _
            "Corel AI"
    End If

End Sub


Sub EnviarPedidoLoteAoAstra()

    Dim pedido As String
    Dim arquivo As Integer
    Dim caminho As String

    pedido = InputBox( _
        "O que você quer fazer com as artes deste lote?", _
        "Corel AI - Produção em Lote")

    If Trim(pedido) = "" Then Exit Sub

    caminho = "C:\CorelAI\pedido_lote.txt"

    arquivo = FreeFile

    Open caminho For Output As #arquivo

    Print #arquivo, pedido

    Close #arquivo

End Sub
Sub AplicarLoteAstra()

    Dim paginaFonte As Page
    Dim sr As ShapeRange
    Dim arte As Shape
    Dim novaPagina As Page

    Dim nomes() As String
    Dim caminho As String
    Dim nomeTemp As String

    Dim i As Long
    Dim j As Long
    Dim total As Long
    Dim encontrou As Boolean

    Set paginaFonte = ActivePage

    If ActiveSelectionRange.Count = 0 Then
        MsgBox _
            "Selecione os grupos das artes do lote.", _
            vbExclamation, _
            "Corel AI"
        Exit Sub
    End If

    Set sr = ActiveSelectionRange
    total = sr.Count

    ReDim nomes(1 To total)

    ' Marca cada arte com um nome temporário estável
    For i = 1 To total

        nomeTemp = "ASTRA_LOTE_" & Format(i, "000")

        sr(i).Name = nomeTemp
        nomes(i) = nomeTemp

    Next i

    ModoLoteAstra = True

    For i = 1 To total

        caminho = "C:\CorelAI\layout_lote_" & _
                  Format(i, "00") & ".txt"

        If Dir(caminho) = "" Then

            ModoLoteAstra = False

            MsgBox _
                "Layout não encontrado:" & vbCrLf & _
                caminho, _
                vbExclamation, _
                "Corel AI"

            Exit Sub

        End If

        ' SEMPRE volta à página original
        paginaFonte.Activate

        Set arte = Nothing
        encontrou = False

        ' Procura novamente a arte pelo nome
        For j = 1 To paginaFonte.Shapes.Count

            If paginaFonte.Shapes(j).Name = nomes(i) Then

                Set arte = paginaFonte.Shapes(j)
                encontrou = True
                Exit For

            End If

        Next j

        If Not encontrou Then

            ModoLoteAstra = False

            MsgBox _
                "Não encontrei a arte " & CStr(i) & _
                " na página original.", _
                vbExclamation, _
                "Corel AI"

            Exit Sub

        End If

        ' Seleciona SOMENTE a arte atual
        arte.CreateSelection

        ' Copia a seleção real
        ActiveSelectionRange.Copy

        ' Cria página exclusiva
        Set novaPagina = ActiveDocument.AddPages(1)

        novaPagina.Activate

        ' Coloca a página na dimensão final ANTES de colar a arte
        AplicarTamanhoPaginaDoLayout novaPagina, caminho

        ' Cola somente esta arte
        novaPagina.ActiveLayer.Paste

        ' Centraliza a arte recém-colada na página
        If ActiveSelectionRange.Count > 0 Then

        ActiveDocument.Unit = cdrMillimeter
        ActiveDocument.ReferencePoint = cdrCenter

        ActiveSelectionRange.SetPosition _
        novaPagina.SizeWidth / 2, _
        novaPagina.SizeHeight / 2

        End If

        ' Aplica o layout na própria página
        NormalizarLayoutLote caminho, i
        AplicarLayoutArquivo caminho, False
        
        Next i

    ModoLoteAstra = False

    paginaFonte.Activate

    MsgBox _
        CStr(total) & " artes processadas.", _
        vbInformation, _
        "Corel AI"

End Sub

Function ObjetoPorIDLoteAstra( _
    ByVal pg As Page, _
    ByVal idAstra As String) As Shape

    Dim partes() As String
    Dim atual As Shape
    Dim raiz As Shape
    Dim i As Long
    Dim indice As Long

    partes = Split(Trim(idAstra), ".")

    If UBound(partes) < 0 Then
        Set ObjetoPorIDLoteAstra = Nothing
        Exit Function
    End If

    ' No lote existe somente UMA arte de topo na página
    If pg.Shapes.Count = 0 Then
        Set ObjetoPorIDLoteAstra = Nothing
        Exit Function
    End If

    Set raiz = pg.Shapes(1)

    ' O ID raiz sempre é 1
    If partes(0) <> "1" Then
        Set ObjetoPorIDLoteAstra = Nothing
        Exit Function
    End If

    Set atual = raiz

    ' Percorre 1.1, 1.2, 1.3.2 etc.
    For i = 1 To UBound(partes)

        If atual.Type <> cdrGroupShape Then
            Set ObjetoPorIDLoteAstra = Nothing
            Exit Function
        End If

        If Not IsNumeric(partes(i)) Then
            Set ObjetoPorIDLoteAstra = Nothing
            Exit Function
        End If

        indice = CLng(partes(i))

        If indice < 1 Or indice > atual.Shapes.Count Then
            Set ObjetoPorIDLoteAstra = Nothing
            Exit Function
        End If

        Set atual = atual.Shapes(indice)

    Next i

    Set ObjetoPorIDLoteAstra = atual

End Function
Sub AplicarTamanhoPaginaDoLayout( _
    ByVal pg As Page, _
    ByVal caminho As String)

    Dim arq As Integer
    Dim linha As String
    Dim partes() As String

    If Dir(caminho) = "" Then Exit Sub

    arq = FreeFile
    Open caminho For Input As #arq

    Do Until EOF(arq)

        Line Input #arq, linha
        linha = Trim(linha)

        If linha <> "" Then

            partes = Split(linha, "|")

            If UCase(Trim(partes(0))) = "PAGE" Then

                If UBound(partes) >= 2 Then

                    ActiveDocument.Unit = cdrMillimeter

                    pg.SetSize _
                        NumeroAstra(partes(1)), _
                        NumeroAstra(partes(2))

                    Exit Do

                End If

            End If

        End If

    Loop

    Close #arq

End Sub

Sub NormalizarLayoutLote( _
    ByVal caminho As String, _
    ByVal numeroArte As Long)

    Dim arq As Integer
    Dim linhas As Collection
    Dim linha As String
    Dim linhaNova As String
    Dim prefixo As String
    Dim i As Long

    If numeroArte <= 1 Then Exit Sub
    If Dir(caminho) = "" Then Exit Sub

    Set linhas = New Collection

    prefixo = CStr(numeroArte) & "."

    arq = FreeFile
    Open caminho For Input As #arq

    Do Until EOF(arq)

        Line Input #arq, linha

        linhaNova = Replace( _
            linha, _
            "|" & prefixo, _
            "|1." _
        )

        linhas.Add linhaNova

    Loop

    Close #arq

    arq = FreeFile
    Open caminho For Output As #arq

    For i = 1 To linhas.Count
        Print #arq, linhas(i)
    Next i

    Close #arq

End Sub

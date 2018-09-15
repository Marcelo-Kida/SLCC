VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubscriber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Componente de Interface para os eventos do COM+

Option Explicit

Implements A6A7A8Publisher.clsPublisher

Implements COMSVCSLib.ObjectControl
Private ObjectContext                       As COMSVCSLib.ObjectContext
Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

'----------------------------------------------------------------------------------------
'Interface para os eventos do COM+
'Controle de Threads ativas
'---------------------------------------------------------------------------------------

Public Sub clsPublisher_AcionaGerenciadores(ByVal pstrParametros As String)

Dim objInterfaceSubscriber                  As A6A7A8VerificaServer.clsInterfaceSubsc
Dim xmlParametro                            As MSXML2.DOMDocument40
Dim strRetorno                              As String
Dim strNomeObjeto                           As String
Dim lngQuantidadeThreadsAtivas              As Long
Dim lngQuantidadeMaxThreads                 As Long
Dim strErro                                 As String

On Error GoTo ErrorHandler

    Set xmlParametro = CreateObject("MSXML2.DOMDocument.4.0")
 
    xmlParametro.loadXML pstrParametros
        
    strNomeObjeto = xmlParametro.selectSingleNode("//NomeObjeto").Text
    
    lngQuantidadeMaxThreads = xmlParametro.selectSingleNode("//QuantidadeMaxThreads").Text
    lngQuantidadeThreadsAtivas = flObterNumeroThreadsAtivos(Trim(strNomeObjeto))
    
    If lngQuantidadeThreadsAtivas <= lngQuantidadeMaxThreads Then
       
       'Controle de Threads
       'Soma um ao numero de threads ativas
       Call flSomarThread(Trim(strNomeObjeto))
       
       Set objInterfaceSubscriber = CreateObject("A6A7A8VerificaServer.clsInterfaceSubsc")
       
       If xmlParametro.selectSingleNode("//NomeFila") Is Nothing Then
           
           If UCase$(xmlParametro.selectSingleNode("//NomeObjeto").Text) = "A8LQS.CLSTREMESSA" Then
                
                strRetorno = objInterfaceSubscriber.LiberarAgendamento(xmlParametro.selectSingleNode("//NomeObjeto").Text)
           
           ElseIf UCase$(xmlParametro.selectSingleNode("//NomeObjeto").Text) = "A8LQS.CLSLIQUIDACAOTQ" Then
           
                strRetorno = objInterfaceSubscriber.LiquidarOperacaoTQ(xmlParametro.selectSingleNode("//NomeObjeto").Text)
                
           ElseIf UCase$(xmlParametro.selectSingleNode("//NomeObjeto").Text) = "A8LQS.CLSPROCESSOMENSAGEMSTR" Then
           
                strRetorno = objInterfaceSubscriber.ProcessaMensagemSTR0008R2Comti(xmlParametro.selectSingleNode("//NomeObjeto").Text)
                            
           ElseIf UCase$(xmlParametro.selectSingleNode("//NomeObjeto").Text) = "A7SERVER.CLSCONTROLEUSUARIOSISTEMA" Then
                
                strRetorno = objInterfaceSubscriber.VerificaUsuariosInativos()
                
            ElseIf UCase$(xmlParametro.selectSingleNode("//NomeObjeto").Text) = "A8LQS.CLSLIQUIDACAOFUTURA" Then
                
                strRetorno = objInterfaceSubscriber.GerenciarFluxoOperacoesLiquidacaoFutura(xmlParametro.selectSingleNode("//NomeObjeto").Text)
                
           End If
           
       Else
          If xmlParametro.selectSingleNode("//Protocolo") Is Nothing Then
       
             strRetorno = objInterfaceSubscriber.ReceberMensagemMQ(xmlParametro.selectSingleNode("//NomeFila").Text, _
                                                                   xmlParametro.selectSingleNode("//NomeObjeto").Text, _
                                                                   xmlParametro.selectSingleNode("//NomeObjetoApoio").Text, _
                                                                   xmlParametro.selectSingleNode("//NomeFilaErro").Text)
          Else
          
              strRetorno = objInterfaceSubscriber.ReceberMensagemMQ(xmlParametro.selectSingleNode("//NomeFila").Text, _
                                                                    xmlParametro.selectSingleNode("//NomeObjeto").Text, _
                                                                    "", _
                                                                    xmlParametro.selectSingleNode("//NomeFilaErro").Text, _
                                                                    xmlParametro.selectSingleNode("//Protocolo").Text)
          End If
       End If
            
       'Controle de Threads
       'Subtrai um ao numero de threads ativas
       Call flSubtrairThread(Trim(strNomeObjeto))
       
       Set objInterfaceSubscriber = Nothing
    Else
        Call flVerificaThreadErro(Trim(strNomeObjeto))
    End If
    
    'Controle de Threads
    'Subtrai um ao numero de threads ativas
    Call flSubtrairThread(Trim(strNomeObjeto))
        
    Set xmlParametro = Nothing

    Exit Sub
ErrorHandler:
    
    strErro = Err.Description

    'Controle de Threads
    'Subtrai um ao numero de threads ativas
    Call flSubtrairThread(Trim(strNomeObjeto))
    
    Set objInterfaceSubscriber = Nothing
    Set xmlParametro = Nothing

    flGravaArquivo strErro

End Sub

'------------------------------------------------------------------------------------------------------
'Gravar os arquivos de erro
'-------------------------------------------------------------------------------------------------------

Private Sub flGravaArquivo(ByVal psErro As String)

Dim strNomeArquivoLogErro                    As String
Dim intFile                                  As Integer

On Error GoTo ErrorHandler

    strNomeArquivoLogErro = App.Path & "/log/" & "Subscriber_" & Format(Now, "yyyymmddHHmmss") & ".log"
    intFile = FreeFile
    Open strNomeArquivoLogErro For Output As intFile
    Print #intFile, psErro
    Close intFile
  
    Exit Sub

ErrorHandler:
    
    Close intFile
        
    Err.Clear
        
End Sub

'------------------------------------------------------------------------------------------------------
'Adicinar uma thread do controle de threads atrivas
'-------------------------------------------------------------------------------------------------------

Private Function flSomarThread(ByVal pstrNomeObjeto As String) As Long

Dim xmlControleThread                       As MSXML2.DOMDocument40
Dim xmlControleThreadAux                    As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleThread                       As String
Dim lngQuantidadeThreadsAtivas              As Long

On Error GoTo ErrorHandler
            
    pstrNomeObjeto = UCase(pstrNomeObjeto)
                
    Set xmlControleThread = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlControleThreadAux = CreateObject("MSXML2.DOMDocument.4.0")
    
    strControleThread = flGetControleThread
        
    'Se o xml de controle de Thread não for gerado
    If strControleThread = vbNullString Then
        
        lngQuantidadeThreadsAtivas = 1
        
        Call fgAppendNode(xmlControleThreadAux, "", "Repeat_Thread", "")
        Call fgAppendNode(xmlControleThreadAux, "Repeat_Thread", "Grupo_Thread", "")
        Call fgAppendAttribute(xmlControleThreadAux, "Grupo_Thread", "NomeObjeto", pstrNomeObjeto)
        Call fgAppendAttribute(xmlControleThreadAux, "Grupo_Thread", "QuantidadeThreadsAtivas", lngQuantidadeThreadsAtivas)
        Call fgAppendAttribute(xmlControleThreadAux, "Grupo_Thread", "UltimaAtualizacao", Format(Now, "yyyymmddHHmmss"))
        Call flSetControleThread(xmlControleThreadAux.xml)
    
    'Se existir o xml de controle de thread
    Else
        
        xmlControleThread.loadXML strControleThread
        
        With xmlControleThread.documentElement
            
            Set xmlNode = .selectSingleNode("Grupo_Thread[@NomeObjeto='" & pstrNomeObjeto & "']")
            
            'Se não existir existir um grupo Thread QuantidadeThreadsAtivas =1
            If xmlNode Is Nothing Then
                
                lngQuantidadeThreadsAtivas = 1
                
                Call fgAppendNode(xmlControleThreadAux, "", "Grupo_Thread", "")
                Call fgAppendAttribute(xmlControleThreadAux, "Grupo_Thread", "NomeObjeto", pstrNomeObjeto)
                Call fgAppendAttribute(xmlControleThreadAux, "Grupo_Thread", "QuantidadeThreadsAtivas", lngQuantidadeThreadsAtivas)
                Call fgAppendAttribute(xmlControleThreadAux, "Grupo_Thread", "UltimaAtualizacao", Format(Now, "yyyymmddHHmmss"))
                Call fgAppendXML(xmlControleThread, "Repeat_Thread", xmlControleThreadAux.xml)

            'Se existir ja existir um um grupo Thread QuantidadeThreadsAtivos = QuantidadeThreadsAtivos + 1
            Else
                
                lngQuantidadeThreadsAtivas = CLng("0" & xmlNode.Attributes.getNamedItem("QuantidadeThreadsAtivas").Text)
                
                lngQuantidadeThreadsAtivas = lngQuantidadeThreadsAtivas + 1
                
                xmlNode.Attributes.getNamedItem("UltimaAtualizacao").Text = Format(Now, "yyyymmddHHmmss")
                xmlNode.Attributes.getNamedItem("QuantidadeThreadsAtivas").Text = lngQuantidadeThreadsAtivas
                                
            End If
            
        End With
        
        Call flSetControleThread(xmlControleThread.xml)
    End If
    
    flSomarThread = lngQuantidadeThreadsAtivas
    
    Set xmlControleThread = Nothing
    Set xmlControleThreadAux = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlControleThread = Nothing
    Set xmlControleThreadAux = Nothing
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'------------------------------------------------------------------------------------------------------
'Subtrair uma thread do controle de threads atrivas
'-------------------------------------------------------------------------------------------------------

Private Function flSubtrairThread(ByVal pstrNomeObjeto As String)

Dim xmlControleThread                       As MSXML2.DOMDocument40
Dim xmlControleThreadAux                    As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strControleThread                       As String
Dim lngQuantidadeThreadsAtivas              As Long

On Error GoTo ErrorHandler
    
    pstrNomeObjeto = Trim(UCase(pstrNomeObjeto))
                
    Set xmlControleThread = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlControleThreadAux = CreateObject("MSXML2.DOMDocument.4.0")
    
    strControleThread = flGetControleThread
        
    'Se o xml de controle de Thread não for gerado sair
    If strControleThread = vbNullString Then
        Exit Function
    Else
        
        xmlControleThread.loadXML strControleThread
        
        With xmlControleThread.documentElement
            
            Set xmlNode = .selectSingleNode("Grupo_Thread[@NomeObjeto='" & pstrNomeObjeto & "']")
            
            'Se não existir existir um grupo Thread QuantidadeThreadsAtivas = 0
            If xmlNode Is Nothing Then
                
                lngQuantidadeThreadsAtivas = 0
                
                Call fgAppendNode(xmlControleThreadAux, "Repeat_Thread", "Grupo_Thread", "")
                Call fgAppendAttribute(xmlControleThreadAux, "Grupo_Thread", "NomeObjeto", pstrNomeObjeto)
                Call fgAppendAttribute(xmlControleThreadAux, "Grupo_Thread", "QuantidadeThreadsAtivas", lngQuantidadeThreadsAtivas)
                
                Call fgAppendXML(xmlControleThread, "Repeat_Thread", xmlControleThreadAux.xml)
            
            'Se existir ja existir um um grupo Thread QuantidadeThreadsAtivos = QuantidadeThreadsAtivos -1
            Else
                
                lngQuantidadeThreadsAtivas = CLng("0" & xmlNode.Attributes.getNamedItem("QuantidadeThreadsAtivas").Text)
                
                lngQuantidadeThreadsAtivas = lngQuantidadeThreadsAtivas - 1
                
                If lngQuantidadeThreadsAtivas = -1 Then lngQuantidadeThreadsAtivas = 0
                
                xmlNode.Attributes.getNamedItem("QuantidadeThreadsAtivas").Text = lngQuantidadeThreadsAtivas
            End If
            
        End With
        
        Call flSetControleThread(xmlControleThread.xml)
    End If
    
    Set xmlControleThread = Nothing
    Set xmlControleThreadAux = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlControleThread = Nothing
    Set xmlControleThreadAux = Nothing
   
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'------------------------------------------------------------------------------------------------------
'Obter o controle de thread gravada no 'Shared Property memory' do COM+
'-------------------------------------------------------------------------------------------------------

Public Function flGetControleThread() As String

Dim objSharedPropMem                        As Object 'A6A7A8.clsSharedPropMem

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    flGetControleThread = objSharedPropMem.GetSPMProperty("A6A7A8", "CONTROLETHREAD")
    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'------------------------------------------------------------------------------------------------------
'Gravar o controle de thread no 'Shared Property memory' do COM+
'-------------------------------------------------------------------------------------------------------

Private Function flSetControleThread(ByVal strControleThread As String) As Boolean

Dim objSharedPropMem                        As Object 'A6A7A8.clsSharedPropMem

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    Call objSharedPropMem.SetSPMProperty("A6A7A8", "CONTROLETHREAD", strControleThread)
    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'------------------------------------------------------------------------------------------------------
'Obter o numero de threads ativas
'-------------------------------------------------------------------------------------------------------


Private Function flObterNumeroThreadsAtivos(ByVal pstrNomeObjeto As String) As Long

Dim objSharedPropMem                        As Object 'A6A7A8.clsSharedPropMem
Dim strControleThread                       As String
Dim xmlControleThread                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
        
    strControleThread = flGetControleThread
    
    If Trim(strControleThread) = vbNullString Then
        flObterNumeroThreadsAtivos = 0
        Exit Function
    End If
        
    Set xmlControleThread = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlControleThread.loadXML strControleThread
    
    With xmlControleThread.documentElement
        
        Set xmlNode = .selectSingleNode("Grupo_Thread[@NomeObjeto='" & UCase(Trim(pstrNomeObjeto)) & "']")
        
        If xmlNode Is Nothing Then
            flObterNumeroThreadsAtivos = 0
        Else
            flObterNumeroThreadsAtivos = CLng("0" & xmlNode.Attributes.getNamedItem("QuantidadeThreadsAtivas").Text)
        End If
        
    End With
    
    Set xmlControleThread = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlControleThread = Nothing
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'------------------------------------------------------------------------------------------------------
'Controle de thread ociosa
'-------------------------------------------------------------------------------------------------------

Private Function flVerificaThreadErro(ByVal pstrNomeObjeto As String) As Long

Dim objSharedPropMem                        As Object 'A6A7A8.clsSharedPropMem
Dim strControleThread                       As String
Dim xmlControleThread                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
        
    strControleThread = flGetControleThread
        
    Set xmlControleThread = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlControleThread.loadXML strControleThread
    
    With xmlControleThread.documentElement
        
        Set xmlNode = .selectSingleNode("Grupo_Thread[@NomeObjeto='" & UCase(Trim(pstrNomeObjeto)) & "']")
        
        If Not xmlNode Is Nothing Then
            If DateDiff("s", fgDtHrStr_To_DateTime(xmlNode.selectSingleNode("@UltimaAtualizacao").Text), Now) > 30 Then
                xmlNode.selectSingleNode("@QuantidadeThreadsAtivas").Text = 0
                xmlNode.selectSingleNode("@UltimaAtualizacao").Text = Format(Now, "yyyymmddHHmmss")
            End If
        End If
        
    End With
        
    Call flSetControleThread(xmlControleThread.xml)
        
    Set xmlControleThread = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlControleThread = Nothing
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function



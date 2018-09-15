VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtrlProcessamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pela manutenção de todos os registros de tabelasd e controle do SLCC
Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

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

'Consultar todas os registros da tabela de controle de processamento de dados - A8.TB_CTRL_PROC_OPER_ATIV
Public Function LerTodos() As String
    
Dim strSQL               As String
    
On Error GoTo ErrorHandler

    strSQL = ""

    strSQL = "SELECT  A.TP_OPER,                     " & vbCrLf & _
             "        B.NO_TIPO_OPER,                " & vbCrLf & _
             "        A.CO_SITU_PROC,                " & vbCrLf & _
             "        C.DE_SITU_PROC,                " & vbCrLf & _
             "        A.NO_PROC_OPER_ATIV,           " & vbCrLf & _
             "        A.TP_LIQU_OPER_ATIV,           " & vbCrLf & _
             "        D.NO_TIPO_LIQU_OPER_ATIV,      " & vbCrLf & _
             "        DECODE(A.IN_ENVI_PREV_PJ,1,'SIM','NAO') AS IN_ENVI_PREV_PJ,               " & vbCrLf & _
             "        DECODE(A.IN_ENVI_PREV_A6,1,'SIM','NAO') AS IN_ENVI_PREV_A6,               " & vbCrLf & _
             "        DECODE(A.IN_ENVI_RELZ_PJ,1,'SIM','NAO') AS IN_ENVI_RELZ_PJ,               " & vbCrLf & _
             "        DECODE(A.IN_ENVI_RELZ_SOLI_A6,1,'SIM','NAO') AS IN_ENVI_RELZ_SOLI_A6,     " & vbCrLf & _
             "        DECODE(A.IN_ENVI_RELZ_CONF_A6,1,'SIM','NAO') AS IN_ENVI_RELZ_CONF_A6,     " & vbCrLf & _
             "        DECODE(A.IN_VERI_REGR_CONF,1,'SIM','NAO') AS IN_VERI_REGR_CONF,           " & vbCrLf & _
             "        DECODE(A.IN_VERI_REGR_CNCL,1,'SIM','NAO') AS IN_VERI_REGR_CNCL,           " & vbCrLf & _
             "        DECODE(A.IN_VERI_REGR_LIBE,1,'SIM','NAO') AS IN_VERI_REGR_LIBE,           " & vbCrLf & _
             "        DECODE(A.IN_ENVI_MESG_RETN,1,'SIM','NAO') AS IN_ENVI_MESG_RETN,           " & vbCrLf & _
             "        DECODE(A.IN_ENVI_MESG_SPB,1,'SIM','NAO') AS IN_ENVI_MESG_SPB,             " & vbCrLf & _
             "        DECODE(A.IN_DISP_LANC_CNTA_CRRT,1,'SIM','NAO') AS IN_DISP_LANC_CNTA_CRRT, " & vbCrLf & _
             "        DECODE(A.IN_ESTO_PJ_A6,1,'SIM','NAO') AS IN_ESTO_PJ_A6,                   " & vbCrLf & _
             "        DECODE(A.IN_ENVI_ALER,1,'SIM','NAO') AS IN_ENVI_ALER                      " & vbCrLf & _
             "FROM    A8.TB_CTRL_PROC_OPER_ATIV   A, " & vbCrLf & _
             "        A8.TB_TIPO_OPER             B, " & vbCrLf & _
             "        A8.TB_SITU_PROC             C, " & vbCrLf & _
             "        A8.TB_TIPO_LIQU_OPER_ATIV D    "
        
    strSQL = strSQL & _
             " WHERE  A.TP_OPER = B.TP_OPER                     " & vbCrLf & _
             " AND    A.CO_SITU_PROC      = C.CO_SITU_PROC      " & vbCrLf & _
             " AND    A.TP_LIQU_OPER_ATIV = D.TP_LIQU_OPER_ATIV " & vbCrLf & _
             " ORDER  BY A.TP_OPER,                             " & vbCrLf & _
             "        A.CO_SITU_PROC,                           " & vbCrLf & _
             "        A.NO_PROC_OPER_ATIV                       "
    
    LerTodos = fgQueryXMLLerTodos("CtrlProcessamento", strSQL, App.EXEName & "." & TypeName(Me))
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function flIncluir(ByRef objDOMDocument As MSXML2.DOMDocument40) As Boolean

Dim strSQL              As String
Dim ExisteMensagem      As Boolean
Dim Rs                  As Recordset

On Error GoTo ErrorHandler
            
    strSQL = " SELECT COUNT(*) AS QTDE " & _
             " FROM   A8.TB_CTRL_PROC_OPER_ATIV " & _
             " WHERE  TP_OPER           = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/TP_OPER").Text) & _
             " AND    CO_SITU_PROC      = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/CO_SITU_PROC").Text) & _
             " AND    TP_LIQU_OPER_ATIV = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/TP_LIQU_OPER_ATIV").Text) & _
             " AND    NO_PROC_OPER_ATIV = '" & Trim(CStr(objDOMDocument.selectSingleNode("//Grupo_Propriedades/NO_PROC_OPER_ATIV").Text)) & "'"

    Set Rs = fgQuerySQL(strSQL)
    
    ExisteMensagem = (Rs!QTDE > 0)
    
    Rs.Close
        
    Set Rs = Nothing
    
    If Not ExisteMensagem Then
           
        strSQL = " INSERT INTO A8.TB_CTRL_PROC_OPER_ATIV " & _
                 "        (TP_OPER,                      " & _
                 "        CO_SITU_PROC,                  " & _
                 "        NO_PROC_OPER_ATIV,             " & _
                 "        IN_ENVI_PREV_PJ,               " & _
                 "        TP_LIQU_OPER_ATIV,             " & _
                 "        IN_ENVI_PREV_A6,               " & _
                 "        IN_ENVI_RELZ_PJ,               " & _
                 "        IN_ENVI_RELZ_SOLI_A6,          " & _
                 "        IN_ENVI_RELZ_CONF_A6,          " & _
                 "        IN_VERI_REGR_CONF,             " & _
                 "        IN_VERI_REGR_CNCL,             " & _
                 "        IN_VERI_REGR_LIBE,             " & _
                 "        IN_ENVI_MESG_RETN,             " & _
                 "        IN_ENVI_MESG_SPB,              " & _
                 "        IN_DISP_LANC_CNTA_CRRT,        " & _
                 "        IN_ESTO_PJ_A6,                 " & _
                 "        IN_ENVI_ALER                   "
    
        strSQL = strSQL & " ) VALUES ( " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/TP_OPER").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/CO_SITU_PROC").Text) & ", " & _
                "'" & Trim(CStr(objDOMDocument.selectSingleNode("//Grupo_Propriedades/NO_PROC_OPER_ATIV").Text)) & "'," & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_PREV_PJ").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/TP_LIQU_OPER_ATIV").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_PREV_A6").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_RELZ_PJ").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_RELZ_SOLI_A6").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_RELZ_CONF_A6").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_VERI_REGR_CONF").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_VERI_REGR_CNCL").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_VERI_REGR_LIBE").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_MESG_RETN").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_MESG_SPB").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_DISP_LANC_CNTA_CRRT").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ESTO_PJ_A6").Text) & ", " & _
                CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_ALER").Text) & ")"
                
        If fgExecuteSQL(strSQL) Then
            flIncluir = True
        Else
            flIncluir = False
        End If
    
    Else
        
        strSQL = " UPDATE A8.TB_CTRL_PROC_OPER_ATIV " & _
                 "    SET IN_ENVI_PREV_PJ        = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_PREV_PJ").Text) & _
                 "    ,   IN_ENVI_PREV_A6        = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_PREV_A6").Text) & _
                 "    ,   IN_ENVI_RELZ_PJ        = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_RELZ_PJ").Text) & _
                 "    ,   IN_ENVI_RELZ_SOLI_A6   = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_RELZ_SOLI_A6").Text) & _
                 "    ,   IN_ENVI_RELZ_CONF_A6   = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_RELZ_CONF_A6").Text) & _
                 "    ,   IN_VERI_REGR_CONF      = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_VERI_REGR_CONF").Text) & _
                 "    ,   IN_VERI_REGR_CNCL      = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_VERI_REGR_CNCL").Text) & _
                 "    ,   IN_VERI_REGR_LIBE      = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_VERI_REGR_LIBE").Text) & _
                 "    ,   IN_ENVI_MESG_RETN      = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_MESG_RETN").Text) & _
                 "    ,   IN_ENVI_MESG_SPB       = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_MESG_SPB").Text) & _
                 "    ,   IN_DISP_LANC_CNTA_CRRT = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_DISP_LANC_CNTA_CRRT").Text) & _
                 "    ,   IN_ESTO_PJ_A6          = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ESTO_PJ_A6").Text) & _
                 "    ,   IN_ENVI_ALER           = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/IN_ENVI_ALER").Text)
               
        strSQL = strSQL & _
                 " WHERE  TP_OPER                = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/TP_OPER").Text) & _
                 " AND    CO_SITU_PROC           = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/CO_SITU_PROC").Text) & _
                 " AND    TP_LIQU_OPER_ATIV      = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/TP_LIQU_OPER_ATIV").Text) & _
                 " AND    NO_PROC_OPER_ATIV      = '" & Trim(CStr(objDOMDocument.selectSingleNode("//Grupo_Propriedades/NO_PROC_OPER_ATIV").Text)) & "'"
               
        If fgExecuteSQL(strSQL) Then
            flIncluir = True
        Else
            flIncluir = False
        End If
    
    End If
      
Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function flExcluir(ByRef objDOMDocument As MSXML2.DOMDocument40) As Boolean

Dim strSQL              As String

On Error GoTo ErrorHandler
            
    strSQL = " DELETE A8.TB_CTRL_PROC_OPER_ATIV " & _
             " WHERE  TP_OPER                = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/TP_OPER").Text) & _
             " AND    CO_SITU_PROC           = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/CO_SITU_PROC").Text) & _
             " AND    TP_LIQU_OPER_ATIV      = " & CLng(objDOMDocument.selectSingleNode("//Grupo_Propriedades/TP_LIQU_OPER_ATIV").Text) & _
             " AND    NO_PROC_OPER_ATIV      = '" & Trim(CStr(objDOMDocument.selectSingleNode("//Grupo_Propriedades/NO_PROC_OPER_ATIV").Text)) & "'"
     
    If fgExecuteSQL(strSQL) Then
        flExcluir = True
    Else
        flExcluir = False
    End If
    
Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


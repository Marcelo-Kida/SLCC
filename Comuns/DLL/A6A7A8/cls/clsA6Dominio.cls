VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA6Dominio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objetivo       : Publicar dominios comuns para os sistemas A6

Option Explicit

Public Enum enumItemCaixaNivel
    Nivel1 = 1
    Nivel2 = 2
    Nivel3 = 3
    Nivel4 = 4
    Nivel5 = 5
End Enum

Public Enum enumStatusProduto
    NaoAssociado = 1
    AssociadoSemFimVigencia = 2
    AssociadoComFimVigencia = 3
End Enum

Public Enum enumSituacaoRemessa
    EmProcessamento = 1
    Finalizado = 2
End Enum

Public Enum enumTipoAbertura
    ValorD1 = 1
    ValorInformado = 2
    SaldoConta = 3
End Enum

'------------------ Definição de Tipo de Mensagem ---------------
Public Enum enumTipoMensagem
    MensagemXML = 1
    MensagemString = 2
    MensagemCSV = 3
    MensagemStringXML = 4
    MensagemCSVXML = 5
End Enum
'------------------ Definição de Status de Caixa ----------------
Public Enum enumEstadoCaixa
    Disponivel = 1
    Aberto = 2
    Fechado = 3
End Enum

Public Enum enumTipoItemCaixa
    Grupo = 1
    Elementar = 2
End Enum

Public Enum enumTipoFiltroA6
    frmSubReservaResumo = 1
    frmControleRemessa = 2
    frmRemessaRejeitada = 3
    frmSubReservaD0 = 4
    frmCaixaFuturo = 5
    frmSubReservaAbertura = 6
    frmSubReservaFechamento = 7
    frmSubReservaConsultaAberturaFechamento = 8
End Enum



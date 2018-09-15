VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsA7Dominio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Objetivo       : Publicar dominios comuns para os sistemas  A7

Option Explicit

Public Enum enumTipoDadoAtributo
    Numerico = 1
    Alfanumerico = 2
End Enum

Public Enum enumOcorrencia
    RecebimentoBemSucedido = 1
    PostagemBemSucedida = 2
    ConfirmacaoEntregaRecebida = 3
    ConfirmacaoRetiradaRecebida = 4
    RejeicaoNaoAutenticidade = 5
    CanceladaErroTradução = 6
    ErroPostagem = 7
    FalhaNaoPrevista = 8
    UsuarioInativo = 9
End Enum

Public Enum enumTipoEntradaMensagem
    EntradaXML = 1
    EntradaString = 2
    EntradaCSV = 3
End Enum

Public Enum enumTipoSaidaMensagem
    NaoseAplica = 0
    SaidaXML = 1
    SaidaString = 2
    SaidaCSV = 3
    SaidaStringXML = 4
    SaidaCSVXML = 5
End Enum

Public Enum enumNaturezaMensagem
    MensagemEnvio = 1
    MensagemConsulta = 2
    MensagemECO = 3
End Enum

Public Enum enumTipoParteSaida
    ParteId = 1
    ParteXML = 2
    ParteSTR = 3
    ParteCSV = 4
    ParteIdLocalLiquidacao = 5
    ParteIdOperacao = 6
End Enum

Public Enum enumVigente
    Vigente = 1
    NaoVigente = 2
End Enum

Public Enum enumRotinaBatch
    ReplicacaoPJPK = 1
    IntegracaoHA = 2
    ExpurgoA6 = 3
    ExpurgoA7 = 4
    ExpurgoA8 = 5
End Enum

Public Enum enumTipoTraducaoBookSPB
    Book2_BookAntigo = 1
    BookAntigo_Book2 = 2
End Enum

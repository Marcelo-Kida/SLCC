VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "clsMIUDominio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa            : Regerbanc - Partticipações , Negócios e Serviços LTDA
'Componente         : BUSMIU
'Classe             : clsMIUDominio
'Data Criação       : 18/07/2003
'Objetivo           :
'                   :
'Analista           : Marcelo Kida
'
'Programador        : Marcelo Kida
'Data               : 18/07/2003
'
'Data Teste         :
'Autor              :
'
'Data Alteração     :
'Autor              :
'Objetivo           :

Option Explicit

Public Enum enumSistemaSLCC
    BUS = 1
    LQS = 2
    SBR = 3
End Enum

Public Enum enumFormatoDataHora
    Data = 1
    Hora = 2
    DataHora = 3
End Enum

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
End Enum

Public Enum enumTipoEntradaMensagem
    EntradaXML = 1
    EntradaString = 2
    EntradaCSV = 3
End Enum

Public Enum enumTipoSaidaMensagem
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
End Enum


Public Enum enumVigente
    Vigente = 1
    NaoVigente = 2
End Enum


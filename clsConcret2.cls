VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConcret2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Buscando a estrutura de abstração
Implements IabstractInterface

Private strOption As String

'Observe que o comportamento das classe muda ou permanece igual.
'Note that the class of behavior changes or stays the same.

'É obrigatório colocar o nome abstrato ("IabstractInterface_") _
 no início dos nomes das Subs, Functions e Propertys!

'It is mandatory to put the abstract name ("IabstractInterface_") _
  At the beginning of Subs, Functions and Property names!

'Veja que aqui tudo é "PRIVATE".
'See that everything here is "PRIVATE".
Private Sub IabstractInterface_Execute()
    Debug.Print strOption
End Sub

Private Property Get IabstractInterface_Opt() As String
    IabstractInterface_Opt = strOption
End Property

Private Property Let IabstractInterface_Opt(strOpt As String)
    strOption = strOpt
End Property

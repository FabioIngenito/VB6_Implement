VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IabstractInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
'Não pode dimensionar variáveis na classe abstrata!
'You can not size variables in the abstract class!
'Dim strOption As String

'Na classe abstrata ficam SOMENTE os cabeçalhos. _
 Somente eles serão aproveitados pelas classe concretas. _
 Isto acontece porque o VB6 é limitado. O VB.NET é bem melhor!

'In the abstract class are ONLY the headers. _
 Only they will be used by the concrete classes. _
 This is because VB6 is limited. VB.NET is much better!

'Veja que aqui tudo é "PUBLIC".
'See that everything here is "PRIVATE".
Public Sub Execute()
End Sub

Public Property Get Opt() As String
End Property

Public Property Let Opt(strOpt As String)
End Property

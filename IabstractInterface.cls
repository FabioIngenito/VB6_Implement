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
'N�o pode dimensionar vari�veis na classe abstrata!
'You can not size variables in the abstract class!
'Dim strOption As String

'Na classe abstrata ficam SOMENTE os cabe�alhos. _
 Somente eles ser�o aproveitados pelas classe concretas. _
 Isto acontece porque o VB6 � limitado. O VB.NET � bem melhor!

'In the abstract class are ONLY the headers. _
 Only they will be used by the concrete classes. _
 This is because VB6 is limited. VB.NET is much better!

'Veja que aqui tudo � "PUBLIC".
'See that everything here is "PRIVATE".
Public Sub Execute()
End Sub

Public Property Get Opt() As String
End Property

Public Property Let Opt(strOpt As String)
End Property

VERSION 5.00
Begin VB.Form frmImplement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Implement"
   ClientHeight    =   1050
   ClientLeft      =   1590
   ClientTop       =   1665
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   3750
   Begin VB.TextBox txtClasse 
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   720
      Width           =   3615
   End
   Begin VB.Frame fraFrame 
      Height          =   675
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   2235
      Begin VB.OptionButton optConcrete 
         Caption         =   "Concreta &2 - Concret 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton optConcrete 
         Caption         =   "Concreta &1 - Concret 1"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdProcessa 
      Caption         =   "&Processa1 - Process"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmImplement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Exemplo bem simples e enxuto do uso do "Implements"
'A very simple and lean example of the use of "Implements"

Private Sub cmdProcessa_Click()
'É dimensionada a classe Abstrata
'The Abstract class is dimensioned
Dim IobjInterface As Implementa.IabstractInterface

    If optConcrete.Item(0).Value Then
        'É instanciada a classe concreta 1
        'Concrete class 1 is instantiated
        Set IobjInterface = New Implementa.clsConcret1
        IobjInterface.Opt = "Concret1"
    Else
        'É instanciada a classe concreta 2
        'Concrete class 2 is instantiated
        Set IobjInterface = New Implementa.clsConcret2
        IobjInterface.Opt = "Concret2"
    End If

    IobjInterface.Execute
    txtClasse.Text = IobjInterface.Opt
    Set IobjInterface = Nothing
End Sub

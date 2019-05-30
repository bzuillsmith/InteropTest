VERSION 5.00
Object = "{4E6B9A16-A197-4FD5-BAB0-43F18AC00781}#1.0#0"; "DotNetControls.tlb"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin DotNetControlsCtl.ExampleButton ExampleButton 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      BackColor       =   "Control"
      ForeColor       =   "ControlText"
      Location        =   "8, 16"
      Name            =   "ExampleButton"
      Size            =   "289, 169"
      Object.TabIndex        =   "0"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExampleButton_Click()
    Dim oTest As MyTestClass
    Set oTest = New MyTestClass
    
    oTest.Message = "test"
    MsgBox oTest.Message
End Sub

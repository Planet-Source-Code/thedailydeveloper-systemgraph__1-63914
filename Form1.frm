VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "myGraph"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin Project1.SystemGraph myGraph 
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   10398
      MaxScale        =   1000
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6600
      Top             =   5880
   End
   Begin VB.CommandButton cmdDrawGraph 
      Caption         =   "Draw Graph"
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private Sub GraphTest()
    Dim i As Integer
    Dim colGraph As New Collection
    
        Do
            
            Randomize
            colGraph.Add Int(Rnd * 1000)
            
            Call myGraph.DrawGraph(colGraph)
            
            tmrWait.Enabled = True
            Do While tmrWait.Enabled = True
                DoEvents
            Loop
            
        Loop
    
    End Sub

    Private Sub cmdDrawGraph_Click()
    
        Call GraphTest
    
    End Sub

    Private Sub Form_Resize()
    Dim intResizeY As Integer
    
        intResizeY = cmdDrawGraph.Height * 2
    
        With myGraph
        
            .Width = Me.Width
            
            ' Only positive values
            If intResizeY < Me.Height Then _
                .Height = Me.Height - intResizeY
        
        End With
        
        With cmdDrawGraph
        
            .Top = (Me.Height - .Height * 2) - 10
            .Width = Me.Width - 100
            
        End With
        
        DoEvents
    
    End Sub

    Private Sub Form_Unload(Cancel As Integer)
    
        End
    
    End Sub

    Private Sub tmrWait_Timer()
    
        tmrWait.Enabled = False
        
    End Sub

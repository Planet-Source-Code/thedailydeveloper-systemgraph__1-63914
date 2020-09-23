VERSION 5.00
Begin VB.UserControl SystemGraph 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   ForeColor       =   &H8000000B&
   HitBehavior     =   0  'None
   ScaleHeight     =   3375
   ScaleWidth      =   5175
   Begin VB.PictureBox Graph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   3375
      Left            =   0
      Picture         =   "SystemGraph.ctx":0000
      ScaleHeight     =   3375
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "SystemGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Max value of graph
Private MaxScaleValue As Integer

    Private Sub UserControl_Initialize()
    
        With Graph
        
            ' Set picturebox cords
             .Left = 0
             .Top = 0
             
             ' Enable autoredraw
             If .AutoRedraw = False Then .AutoRedraw = True
         
        End With

    End Sub

    Private Sub UserControl_Resize()
    
        With Graph
        
            ' Rezize control
            .Width = UserControl.Width
            .Height = UserControl.Height
        
        End With
        
    End Sub
    
    Public Function DrawGraph(ByVal colGraphValues As Collection)
    Dim i As Long
    Dim IntX(1) As Integer
    Dim IntY(1) As Integer
    Dim intYScale As Integer
    
        ' Just continue
        On Local Error Resume Next
        
        With Graph
    
            ' Clear graph
            .Cls
            
            ' Calculate scale
            intYScale = .ScaleHeight / MaxScaleValue
            
            ' Keep the collection inside the bounarys of our graph
            Do While colGraphValues.Count > Int(.Width / 100 / 2) 'ADDED /2
            
                ' Remove oldest
                colGraphValues.Remove 1
                
            Loop
        
            ' Set x-cord
            .CurrentX = 0
        
            ' Draw all y- scales exept the zero
            For i = 0 To 4
                
                ' Draw scales
                .CurrentY = Int(.ScaleHeight / 4) * i
                Graph.Print MaxScaleValue - Int(MaxScaleValue * i / 4)
                
            Next i
            
            ' Print y- scale 0
            .CurrentY = .ScaleHeight - 200
            Graph.Print 0
            
        End With
        
        Graph.ForeColor = &HC0C0C0
        
        ' Draw lines
        For i = 1 To colGraphValues.Count
        
            ' Set new
            IntX(1) = IntX(0) + 200
            IntY(1) = colGraphValues(i) * intYScale
        
            ' Draw lines
            Graph.Line (Graph.Width - IntX(0), Graph.Height - IntY(0))-(Graph.Width - IntX(1), Graph.Height - IntY(1)), vbGreen
        
            ' Set old
            IntX(0) = IntX(1)
            IntY(0) = IntY(1)
        
        Next i
        DoEvents
        
    End Function

    Public Property Get MaxScale() As Integer
    
        ' Read maxscale
        MaxScale = MaxScaleValue
        
    End Property
    
    Public Property Let MaxScale(intScale As Integer)
        
        ' Set maxscale
        MaxScaleValue = intScale
        
        ' Tell usercontrol
        PropertyChanged "MaxScale"
        
    End Property
    
    Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
        ' Read maxscale property from usercontrol
        MaxScaleValue = PropBag.ReadProperty("MaxScale", 100)
            
    End Sub
    
    Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
        ' Write maxscale property to usercontrol
        Call PropBag.WriteProperty("MaxScale", MaxScaleValue, 100)
        
    End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2205
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   2205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton EndPolygon 
      Caption         =   "End Polygon"
      Height          =   585
      Left            =   30
      TabIndex        =   2
      Top             =   2610
      Width           =   2145
   End
   Begin VB.CommandButton BeginPolygon 
      Caption         =   "Begin Polygon"
      Height          =   585
      Left            =   30
      TabIndex        =   1
      Top             =   1980
      Width           =   2145
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   30
      ScaleHeight     =   1875
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   30
      Width           =   2145
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ScreenPoints As Collection
Public ScreenPoint As clsPoint
Private PlotInProgress As Boolean
Private PolygonFinished As Boolean

' runs too slow...use for testing
'Private Sub FillPolygon_Click()
'    Dim W As Single
'    Dim H As Single
'    Dim Target As clsPoint
'
'    For W = 100 To Picture1.Width - 100
'        For H = 100 To Picture1.Height - 100
'            If (PolygonFinished) Then
'                Set Target = New clsPoint
'                Target.x = W
'                Target.y = H
'                If (PtInPolygon(ScreenPoints, Target)) Then
'                    Picture1.PSet (W, H), vbBlue
'                Else
'                    Picture1.PSet (W, H), vbRed
'                End If
'            End If
'        Next
'    Next
'End Sub

Private Sub Form_Load()
    Call Init
End Sub

Private Sub BeginPolygon_Click()
    Call Init
End Sub

Private Sub EndPolygon_Click()
    Dim FirstPoint As clsPoint
    
    Set FirstPoint = ScreenPoints.Item(1)
    Call ProcessClick(FirstPoint.x, FirstPoint.y)
    PolygonFinished = True
    EndPolygon.Enabled = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = MouseButtonConstants.vbLeftButton) Then
        If Not (PlotInProgress) Then
            PlotInProgress = True
        End If
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Target As clsPoint

    If (PolygonFinished) Then
        
        Set Target = New clsPoint
        Target.x = x
        Target.y = y
        If (PtInPolygon(ScreenPoints, Target)) Then
            Picture1.PSet (x, y), vbBlack
        End If
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = MouseButtonConstants.vbLeftButton) Then
        If (PlotInProgress) Then
            PlotInProgress = False
            If Not (PolygonFinished) Then
                Call ProcessClick(x, y)
            End If
        End If
    End If
End Sub

Private Sub ProcessClick(ByVal x As Single, ByVal y As Single)
    Dim LastPoint As clsPoint

    Set ScreenPoint = New clsPoint
    ScreenPoint.x = x
    ScreenPoint.y = y
    Call ScreenPoints.Add(ScreenPoint)
    Picture1.PSet (ScreenPoint.x, ScreenPoint.y), vbBlue
        
    If (ScreenPoints.Count > 1) Then
        Set LastPoint = ScreenPoints.Item(ScreenPoints.Count - 1)
        Picture1.Line (LastPoint.x, LastPoint.y)-(ScreenPoint.x, ScreenPoint.y), vbBlack
        If (ScreenPoints.Count > 2) Then
            EndPolygon.Enabled = True
        End If
    End If
End Sub

Private Sub Init()
    PlotInProgress = False
    Set ScreenPoints = New Collection
    Call Picture1.Cls
    PolygonFinished = False
    EndPolygon.Enabled = False
End Sub

Public Function PtInPolygon(ByRef Points As Collection, ByRef Target As clsPoint) As Boolean
    Dim Source1 As clsPoint
    Dim Source2 As clsPoint
    Dim Index As Integer
    Dim OddNodes As Boolean

    OddNodes = False
    For Index = 1 To Points.Count
        Set Source1 = Points.Item(Index)
        Set Source2 = Points.Item(IIf((Index + 1) > Points.Count, 1, Index + 1))
        If (Source1.y < Target.y And Source2.y >= Target.y Or Source2.y < Target.y And Source1.y >= Target.y) Then
            If (Source1.x + (Target.y - Source1.y) / (Source2.y - Source1.y) * (Source2.x - Source1.x) < Target.x) Then
                OddNodes = Not OddNodes
            End If
        End If
    Next Index
    PtInPolygon = OddNodes
End Function

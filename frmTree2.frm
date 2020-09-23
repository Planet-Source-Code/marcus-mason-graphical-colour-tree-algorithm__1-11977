VERSION 5.00
Begin VB.Form frmTree 
   AutoRedraw      =   -1  'True
   Caption         =   "Tree"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   DrawWidth       =   2
   FillColor       =   &H000080FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTree 
      AutoRedraw      =   -1  'True
      Height          =   6015
      Left            =   4320
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   21
      Top             =   120
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtColourChange 
         Height          =   285
         Left            =   2760
         TabIndex        =   20
         Text            =   "+20"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox txtStartColour 
         Height          =   285
         Left            =   2760
         TabIndex        =   18
         Text            =   "128"
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txtThicknessReduction 
         Height          =   285
         Left            =   2760
         TabIndex        =   16
         Text            =   "0.7"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox txtStartThickness 
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Text            =   "10"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtMinBranchSize 
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Text            =   "3"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtForkAngle 
         Height          =   285
         Left            =   2760
         TabIndex        =   13
         Text            =   "50"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtRandomRadius 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Text            =   "0.2"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtBranchReduction 
         Height          =   285
         Left            =   2760
         TabIndex        =   11
         Text            =   "0.65"
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtStartLength 
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Text            =   "140"
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "Draw"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Tree Code and Alogirthm by Marcus Mason.  Email:M.J.Mason@cs.cf.ac.uk www.planet-source-code.com"
         Height          =   615
         Left            =   240
         TabIndex        =   22
         Top             =   4680
         Width           =   3615
      End
      Begin VB.Label Label9 
         Caption         =   "Colour Change"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Start Colour"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Branch Thickness Reduction"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Start Branch Thickness"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Branch Start Length"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Branch Reduction"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Branch Randomness"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Fork Angle"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Min branch size"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Const Pi = 3.1415926
Dim MinBranchSize As Integer
Dim ForkAngle As Integer
Dim BranchReduction As Single
Dim ThicknessReduction As Single
Dim RandomRadius As Single
Dim ColourChange As Integer

Sub DrawTree(ByVal X As Integer, ByVal Y As Integer, ByVal StartLength As Integer, ByVal StartThickness As Integer, ByVal StartColour As Integer)
    
    'Manually draw trunk of tree and initiate branch algorithm
    
    'Trunk thickness
    picTree.DrawWidth = StartThickness
    
    'Trunk colour
    picTree.ForeColor = RGB(0, StartColour, 0)
    
    'Draw trunk
    picTree.Line (X, Y)-(X, Y - StartLength)
    
    'Call branch algorithm
    Branch X, Y - StartLength, StartLength * BranchReduction, 90, StartThickness * ThicknessReduction, StartColour + ColourChange

End Sub

Sub Branch(ByVal X As Integer, ByVal Y As Integer, ByVal Length As Integer, ByVal Angle As Integer, ByVal Thickness As Integer, ByVal Colour As Integer)
    'Inputs:
    '
    'x,y - coordinates of branch point
    'length - length of the previous branch
    'angle - angle of the last branch
    '       (angles start from 3 O'clock and go around anticlockwise)
    'thickness - thickness of the previous branch
    'colour - the colour of the last branch
    '
    'Function: given a point will draw a fork of
    '           two branches divergent from that point.
    '           A call is made at either end of the branches to
    '           draw further branches.
    '           Algorithm cancels itself when the branches is too small
    
    
    Dim X1 As Integer   '
    Dim Y1 As Integer   'Coordinates of end of branch1
    Dim X2 As Integer   '
    Dim Y2 As Integer   'Coordinates of end of branch2
    Dim Ang1 As Double  'Angle of branch 1
    Dim Ang2 As Double  'Angle of branch 2
    
    'Check if branch length is too small
    If Length < MinBranchSize Then Exit Sub
    
    'Calculate the angles of the diverging branches
    Ang1 = CRad(Angle + 0.5 * ForkAngle)
    Ang2 = CRad(Angle - 0.5 * ForkAngle)
    
    'Calculate the coordinates of the ends of the branches
    'based on the angle they protrude at and the length of the branch
    X1 = X + (Length * Cos(Ang1))
    Y1 = Y - (Length * Sin(Ang1))
    
    X2 = X + (Length * Cos(Ang2))
    Y2 = Y - (Length * Sin(Ang2))
    
    'Add some noise to the branch coordinates, related to length of branch
    RandomVariation X1, Y1, Length * RandomRadius
    RandomVariation X2, Y2, Length * RandomRadius
    
    'Check thickness is not too small
    If Thickness = 0 Then Thickness = 1
    
    'Set branch thickness
    picTree.DrawWidth = Thickness
    
    'Check colour value is not out of range
    If Colour < 0 Then Colour = 0
    If Colour > 255 Then Colour = 255
    
    'Draw the branches
    picTree.Line (X, Y)-(X1, Y1), RGB(0, Colour, 0)
    picTree.Line (X, Y)-(X2, Y2), RGB(0, Colour, 0)
    
    'Call the next branches
    Branch X1, Y1, Length * BranchReduction, CDeg(Ang1), Thickness * ThicknessReduction, Colour + ColourChange
    Branch X2, Y2, Length * BranchReduction, CDeg(Ang2), Thickness * ThicknessReduction, Colour + ColourChange
    
End Sub

Function CRad(ByVal Deg As Integer)
    'convert deg to rad
    CRad = Deg * Pi / 180
End Function
Function CDeg(ByVal Rad As Double)
    'Convert rad to deg
    CDeg = Rad * 180 / Pi
End Function

Function RandomVariation(ByRef X As Integer, ByRef Y As Integer, ByVal Radius As Integer)
    'Input: coordinates X,Y and radius of error
    '
    'Output: coordinates X,Y with randomness
    '
    'Function: add randomness which is evenly distributed around a point X,Y,
    '           the max radius of error is RandomRadiues
    Dim RndRadius As Single
    Dim RndAngle As Single
    
    
    'Calculate a random radius
    RndRadius = Rnd() * Radius
    
    'Calculate a random angle
    RndAngle = CRad(Int(Rnd() * 360))
    
    'Add randomness to point
    X = X + RndRadius * Cos(RndAngle)
    Y = Y + RndRadius * Sin(RndAngle)
End Function



Private Sub cmdClear_Click()
    'Clear the picture box
    picTree.Cls
End Sub

Private Sub cmdDraw_Click()
    'Get user options
    MinBranchSize = Val(txtMinBranchSize)
    ForkAngle = Val(txtForkAngle)
    BranchReduction = Val(txtBranchReduction)
    RandomRadius = Val(txtRandomRadius)
    ThicknessReduction = Val(txtThicknessReduction)
    ColourChange = Val(txtColourChange)
    
    'Clear the piccy box
    picTree.Cls
    
    'Initiate draw of tree
    DrawTree picTree.Width / 30, (picTree.Height / 15) - 20, Val(txtStartLength), Val(txtStartThickness), Val(txtStartColour)
End Sub


Private Sub Form_Resize()
    'Resize pic box to fit the form
    picTree.Width = Me.Width - 4590
    picTree.Height = Me.Height - 630
End Sub

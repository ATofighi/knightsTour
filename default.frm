VERSION 5.00
Begin VB.Form theForm 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Õ—ò  „Â—Â «”»"
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "default.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "default.frx":164A
   RightToLeft     =   -1  'True
   ScaleHeight     =   522
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton resetButton 
      Caption         =   "»«“‰‘«‰Ì"
      Height          =   285
      Left            =   540
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   7335
      Width           =   870
   End
   Begin VB.ComboBox speedList 
      Height          =   315
      ItemData        =   "default.frx":EEA9
      Left            =   4140
      List            =   "default.frx":EEC8
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   64
      Top             =   7335
      Width           =   1785
   End
   Begin VB.Timer horseAnimateTimer 
      Interval        =   16
      Left            =   6675
      Top             =   1920
   End
   Begin VB.Timer moveTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6675
      Top             =   1395
   End
   Begin VB.Label closeButton 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "◊"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   135
      MouseIcon       =   "default.frx":EEF5
      MousePointer    =   99  'Custom
      RightToLeft     =   -1  'True
      TabIndex        =   68
      ToolTipText     =   "»” ‰"
      Top             =   45
      Width           =   675
   End
   Begin VB.Label alaki 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2340
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   7380
      Width           =   45
   End
   Begin VB.Label speedLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "”—⁄ :"
      Height          =   195
      Left            =   5970
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   7380
      Width           =   555
   End
   Begin VB.Image theHeader 
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   7200
   End
   Begin VB.Image chessLoading 
      Height          =   6030
      Left            =   -6240
      MousePointer    =   11  'Hourglass
      Picture         =   "default.frx":F047
      Top             =   885
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Image horse 
      Height          =   750
      Left            =   5130
      Picture         =   "default.frx":2A9D1
      Top             =   1065
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   63
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   62
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   61
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   60
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   59
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   58
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   57
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   56
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   55
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   54
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   53
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   52
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   51
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   50
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   49
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   48
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   47
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   46
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   45
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   44
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   43
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   42
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   41
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   40
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   39
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   38
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   37
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   36
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   35
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   34
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   33
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   32
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   31
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   30
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   29
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   28
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   27
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   26
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   25
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   24
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   23
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   22
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   21
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   20
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   19
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   18
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   17
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   16
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   15
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   14
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   13
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   12
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   11
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   10
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   9
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   8
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   7
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   6
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   5
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   4
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   3
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   2
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   0
      Left            =   0
      MouseIcon       =   "default.frx":2B0CE
      MousePointer    =   99  'Custom
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   45
   End
   Begin VB.Label cell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   1
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   45
   End
   Begin VB.Image chessBackground 
      Height          =   6030
      Left            =   555
      Picture         =   "default.frx":2B220
      Top             =   1230
      Width           =   6000
   End
   Begin VB.Image inTheNameOfGod 
      Height          =   675
      Left            =   330
      Picture         =   "default.frx":4EA50
      Top             =   465
      Width           =   6750
   End
End
Attribute VB_Name = "theForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' In the name of GOD
Const chessSize = 8
Dim mark(chessSize, chessSize) As Integer
Dim ways(chessSize, chessSize) As Integer
Dim moves(chessSize * chessSize) As Integer
Dim moveCount1, moveCount2 As Integer
Dim horseLeftSpeed, horseTopSpeed, horseLeft, horseTop As Integer
Dim runned As Boolean
Dim closeButtonIsHover As Boolean

Dim FrmMove, DragX, DragY ' top bar dragable

' closeButton hover
Private Sub closeButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If closeButtonIsHover = False Then closeButton.BackColor = RGB(235, 80, 80)
    closeButtonIsHover = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If closeButtonIsHover = True Then closeButton.BackColor = RGB(199, 80, 80)
    closeButtonIsHover = False
End Sub


Private Sub resetButton_Click()
    horse.Visible = False
    moveTimer.Enabled = False
    chessLoading.Visible = False
    moveCount1 = 0
    moveCount2 = 0
    For i = 0 To (chessSize - 1)
        For j = 0 To (chessSize - 1)
            mark(i, j) = 0
            cell(k).Caption = ""
        Next j
    Next i
    runned = False

    For i = 0 To (chessSize - 1)
        For j = 0 To (chessSize - 1)
            k = xyToNum(i, j)
            cell(k).MousePointer = 99
            cell(k).MouseIcon = closeButton.MouseIcon
            cell(k).Caption = ""
            X = 0
            X = X + isOk(i + 2, j + 1)
            X = X + isOk(i + 2, j - 1)
            X = X + isOk(i - 2, j + 1)
            X = X + isOk(i - 2, j - 1)
            X = X + isOk(i + 1, j + 2)
            X = X + isOk(i + 1, j - 2)
            X = X + isOk(i - 1, j + 2)
            X = X + isOk(i - 1, j - 2)
            ways(i, j) = X
        Next j
    Next i
End Sub

Private Sub speedList_Click()
    moveTimer.Interval = speedList.ItemData(speedList.ListIndex)
End Sub

Private Sub speedList_Change()
    moveTimer.Interval = speedList.ItemData(speedList.ListIndex)
End Sub


' topBar dragable

Private Sub theHeader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FrmMove = True
    DragX = X
    DragY = Y
End Sub
Private Sub theHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'closeButton hover
    If closeButtonIsHover = True Then closeButton.BackColor = RGB(199, 80, 80)
    closeButtonIsHover = False

    
    Dim nx, ny
    If FrmMove = True Then
        nx = theForm.Left + X - DragX
        ny = theForm.Top + Y - DragY
        'bachasbe be bala|chap
        If nx < 0 Then nx = 0
        If ny < 0 Then ny = 0
        'bechasbe be payin|rast
        If nx + theForm.Width > Screen.Width Then nx = Screen.Width - theForm.Width
        If ny + theForm.Height > Screen.Height Then ny = Screen.Height - theForm.Height
        theForm.Left = nx
        theForm.Top = ny
    End If
End Sub
Public Sub theHeader_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nx, ny
    nx = theForm.Left + X - DragX
    ny = theForm.Top + Y - DragY
    If nx < 0 Then nx = 0
    If ny < 0 Then ny = 0
    If nx + theForm.Width > Screen.Width Then nx = Screen.Width - theForm.Width
    If ny + theForm.Height > Screen.Height Then ny = Screen.Height - theForm.Height
    theForm.Left = nx
    theForm.Top = ny
    FrmMove = False
End Sub





Function xyToNum(X, Y)
    xyToNum = X * chessSize + Y
End Function

Function numToX(X)
    numToX = Int(X / chessSize)
End Function

Function numToY(X)
    numToY = X Mod chessSize
End Function

Function isOk(X, Y)
    If (X < 0 Or X > (chessSize - 1) Or Y < 0 Or Y > (chessSize - 1)) Then
        isOk = 0
    ElseIf mark(X, Y) > 0 Then
        isOk = 0
    Else
        isOk = 1
    End If
End Function

Function runMove(X, Y)
    If isOk(X, Y) = False Then
        runMove = False
    Else
        moves(moveCount1) = xyToNum(X, Y)
        moveCount1 = moveCount1 + 1
        mark(X, Y) = moveCount1
        nextX = -1
        nextY = -1
        nextWays = -1
        xp = X + 2: yp = Y + 1: If isOk(xp, yp) Then ways(xp, yp) = ways(xp, yp) - 1: If ways(xp, yp) < nextWays Or nextWays = -1 Then nextWays = ways(xp, yp): nextX = xp: nextY = yp
        xp = X + 2: yp = Y - 1: If isOk(xp, yp) Then ways(xp, yp) = ways(xp, yp) - 1: If ways(xp, yp) < nextWays Or nextWays = -1 Then nextWays = ways(xp, yp): nextX = xp: nextY = yp
        xp = X - 2: yp = Y + 1: If isOk(xp, yp) Then ways(xp, yp) = ways(xp, yp) - 1: If ways(xp, yp) < nextWays Or nextWays = -1 Then nextWays = ways(xp, yp): nextX = xp: nextY = yp
        xp = X - 2: yp = Y - 1: If isOk(xp, yp) Then ways(xp, yp) = ways(xp, yp) - 1: If ways(xp, yp) < nextWays Or nextWays = -1 Then nextWays = ways(xp, yp): nextX = xp: nextY = yp
        xp = X + 1: yp = Y + 2: If isOk(xp, yp) Then ways(xp, yp) = ways(xp, yp) - 1: If ways(xp, yp) < nextWays Or nextWays = -1 Then nextWays = ways(xp, yp): nextX = xp: nextY = yp
        xp = X + 1: yp = Y - 2: If isOk(xp, yp) Then ways(xp, yp) = ways(xp, yp) - 1: If ways(xp, yp) < nextWays Or nextWays = -1 Then nextWays = ways(xp, yp): nextX = xp: nextY = yp
        xp = X - 1: yp = Y + 2: If isOk(xp, yp) Then ways(xp, yp) = ways(xp, yp) - 1: If ways(xp, yp) < nextWays Or nextWays = -1 Then nextWays = ways(xp, yp): nextX = xp: nextY = yp
        xp = X - 1: yp = Y - 2: If isOk(xp, yp) Then ways(xp, yp) = ways(xp, yp) - 1: If ways(xp, yp) < nextWays Or nextWays = -1 Then nextWays = ways(xp, yp): nextX = xp: nextY = yp
        r = runMove(nextX, nextY)
        runMove = r
    End If
End Function

Private Sub closeButton_Click()
    End
End Sub

Private Sub Form_Load()
    closeButton.BackColor = RGB(199, 80, 80)
    speedList.ListIndex = 2
    runned = False
    chessLoading.Left = chessBackground.Left
    chessLoading.Top = chessBackground.Top

    For i = 0 To (chessSize - 1)
        For j = 0 To (chessSize - 1)
            k = xyToNum(i, j)
            cell(k).MouseIcon = closeButton.MouseIcon
            cell(k).MousePointer = 99
            cell(k).AutoSize = False
            cell(k).Width = 50
            cell(k).Height = 50
            cell(k).Alignment = 2
            cell(k).Top = chessBackground.Top + i * 50
            cell(k).Left = chessBackground.Left + j * 50
            cell(k).FontSize = 28
            If i Mod 2 <> j Mod 2 Then
                cell(k).ForeColor = RGB(255, 255, 255)
            End If
            X = 0
            X = X + isOk(i + 2, j + 1)
            X = X + isOk(i + 2, j - 1)
            X = X + isOk(i - 2, j + 1)
            X = X + isOk(i - 2, j - 1)
            X = X + isOk(i + 1, j + 2)
            X = X + isOk(i + 1, j - 2)
            X = X + isOk(i - 1, j + 2)
            X = X + isOk(i - 1, j - 2)
            'MsgBox (x)
            ways(i, j) = X
            'cell(k).Caption = x
        Next j
    Next i
End Sub

Private Sub Cell_click(cellNumber As Integer)
    If runned = False Then
        X = numToX(cellNumber)
        Y = numToY(cellNumber)
        newX = X
        newY = Y
        If X < (chessSize / 2) Then newX = (chessSize - 1) - X
        If Y < (chessSize / 2) Then newY = (chessSize - 1) - Y
        chessLoading.Visible = True
        r = runMove(newX, newY)
        For i = 0 To (chessSize * chessSize - 1)
            cell(i).MousePointer = 0
            theX = numToX(moves(i))
            theY = numToY(moves(i))
            If X < 4 Then theX = (chessSize - 1) - theX
            If Y < 4 Then theY = (chessSize - 1) - theY
            moves(i) = xyToNum(theX, theY)
        Next i
        moveTimer.Enabled = True
    End If
    runned = True
End Sub


Private Sub horseAnimateTimer_Timer()
    If horseLeftSpeed <> 0 Then
        If (horseLeft - horse.Left) * (horseLeft - (horse.Left + horseLeftSpeed)) < 0 Then
            horse.Left = horseLeft
        Else
            horse.Left = horse.Left + horseLeftSpeed
        End If
    End If
    If horseTopSpeed <> 0 Then
        If (horseTop - horse.Top) * (horseTop - (horse.Top + horseTopSpeed)) < 0 Then
            horse.Top = horseTop
        Else
            horse.Top = horse.Top + horseTopSpeed
        End If
    End If
    If horse.Left = horseLeft Then horseLeftSpeed = 0
    If horse.Top = horseTop Then horseTopSpeed = 0
    
End Sub

Private Sub moveTimer_Timer()
    chessLoading.Visible = False
    cellNumber = moves(moveCount2)
    moveCount2 = moveCount2 + 1
    If moveCount2 = 1 Then
        horse.Left = cell(cellNumber).Left
        horse.Top = cell(cellNumber).Top
    Else
        cell(moves(moveCount2 - 2)).Caption = moveCount2 - 1
        horseLeft = cell(cellNumber).Left
        horseTop = cell(cellNumber).Top
        disLeft = Abs(horseLeft - horse.Left)
        disTop = Abs(horseTop - horse.Top)
        horseLeftSpeed = (40 * (horseLeft - horse.Left) / (moveTimer.Interval))
        horseTopSpeed = (40 * (horseTop - horse.Top) / (moveTimer.Interval))
    End If
    horse.Visible = True
    If moveCount2 > (chessSize * chessSize - 1) Then
        moveTimer.Enabled = False
    End If
End Sub

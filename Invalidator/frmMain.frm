VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invalidator"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   120
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picCtrls 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   2520
      Width           =   6015
      Begin VB.TextBox txtFileName 
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   0
         Width           =   5055
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open"
         Height          =   495
         Left            =   840
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdInvalidate 
         Caption         =   "&Invalidate"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   495
         Left            =   4680
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   705
      End
   End
   Begin VB.Label lblCaption 
      Caption         =   $"frmMain.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Label lblCaption 
      Caption         =   $"frmMain.frx":00B1
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1095
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status ="
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   6285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdInvalidate_Click()
    picCtrls.Enabled = False: DoEvents
    
    Dim bytData() As Byte
    Dim TopByteID As Long: TopByteID = 1023
    ReDim bytData(TopByteID) As Byte
    Dim lLoc As Long
    Dim lLof As Long
    
    Dim ByteCount As Integer
    Open txtFileName.Text For Binary Access Read Write As #1
        If Err.Number <> 0 Then
            MsgBox "File not found or read-only.", vbCritical
            picCtrls.Enabled = True: DoEvents
            Exit Sub
        End If
        
        lLoc = 1
        lLof = LOF(1)
        Do While lLoc + 1024 <= lLof
            Get #1, , bytData
            For ByteCount = 0 To TopByteID
                bytData(ByteCount) = Not bytData(ByteCount)
            Next
            Seek #1, lLoc
            lblStatus.Caption = "Status = " & lLoc & " of " & lLof - 1 & " bytes Invalidating.": DoEvents
            Put #1, , bytData
            lLoc = lLoc + 1024
        Loop
        
        TopByteID = lLof - lLoc
        If TopByteID > 0 Then
            TopByteID = TopByteID - 1
            ReDim bytData(TopByteID)
            Get #1, , bytData
            For ByteCount = 0 To TopByteID
                bytData(ByteCount) = Not bytData(ByteCount)
            Next
            Seek #1, lLoc
            lblStatus.Caption = "Status = " & lLoc & " of " & lLof - 1 & " bytes Invalidating.": DoEvents
            Put #1, , bytData
            If Loc(1) + 1 = LOF(1) Then
                lblStatus.Caption = "Status = Invalidation Successfully done."
            Else
                lblStatus.Caption = "Status = Programming Calculation Mistakes found."
            End If
        ElseIf TopByteID = 0 Then
            lblStatus.Caption = "Status = Invalidation Successfully done."
        Else
            lblStatus.Caption = "The file is damaged. Please try again to repear."
        End If
    Close #1
    picCtrls.Enabled = True
End Sub

Private Sub cmdOpen_Click()
    Dim FileName As String
    
    On Error Resume Next
    With cdlOpen
        .CancelError = True
        .Filter = "All Files (*.*)|*.*"
        .Flags = cdlOFNFileMustExist
        .ShowOpen
        If Err.Number = cdlCancel Then Exit Sub
        txtFileName.Text = .FileName
    End With
End Sub

Private Sub txtFileName_Change()
    On Error Resume Next
    lblStatus.Caption = "Status = " & FileLen(txtFileName.Text) & " Bytes rady to Invalidate."
End Sub

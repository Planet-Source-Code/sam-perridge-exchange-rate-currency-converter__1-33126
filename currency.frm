VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Currency 
   Caption         =   "Currency"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   2955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDExchange 
      Caption         =   "Exchange"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtValue 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton CMDRates 
      Caption         =   "Get Rates"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   2655
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Currency (eg. $25.0)"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Currency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UK As Single
Dim money As Single
Private Sub CMDExchange_Click()
If UK = 0 Then
    CMDRates_Click
End If
If (InStr(1, txtValue, "$") <> 0) Then
    txtValue = Right(txtValue, Len(txtValue) - 1)
    money = Trim(txtValue)
    txtValue = money * UK
    txtValue = "£" & txtValue
ElseIf (InStr(1, txtValue, "£") <> 0) Then
    txtValue = Right(txtValue, Len(txtValue) - 1)
    money = Trim(txtValue)
    txtValue = money / UK
    txtValue = "$" & txtValue
Else
    MsgBox "Please add $ or £ before the value"
End If

End Sub

Private Sub CMDRates_Click()
Dim UKstart As Long
Dim UKend As Long
Dim PageHtml As String
PageHtml = Inet.OpenURL("http://uk.moneycentral.msn.com/investor/market/rates.asp")
UKstart = InStr(1, PageHtml, "US dollar&nbsp;&nbsp;&nbsp;</TD><TD ALIGN=RIGHT>") + Len("US dollar&nbsp;&nbsp;&nbsp;</TD><TD ALIGN=RIGHT>")
UKend = InStr(UKstart, PageHtml, "</TD>")
UK = Mid(PageHtml, UKstart, UKend - UKstart)
MsgBox UK & " pounds to the dollar"
End Sub

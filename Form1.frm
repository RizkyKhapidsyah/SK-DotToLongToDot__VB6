VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Network Addresses Numbers"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1635
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "http://www.cis.ohio-state.edu/cgi-bin/rfc/rfc0943.html"
      Top             =   6180
      Width           =   4800
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1635
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "http://www.cis.ohio-state.edu/cgi-bin/rfc/rfc0952.html"
      Top             =   5820
      Width           =   4800
   End
   Begin VB.Frame Frame2 
      Caption         =   "The Way We Use Them In Windows Programming"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   90
      TabIndex        =   12
      Top             =   15
      Width           =   8535
      Begin VB.CommandButton Command1 
         Caption         =   """Long"" To Dot IP Address"
         Height          =   375
         Left            =   90
         TabIndex        =   17
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2220
         TabIndex        =   16
         Top             =   495
         Width           =   2520
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2220
         TabIndex        =   15
         Top             =   1155
         Width           =   2520
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Dot IP Address To ""Long"""
         Height          =   375
         Left            =   105
         TabIndex        =   14
         Top             =   1140
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   330
         TabIndex        =   13
         Top             =   2160
         Width           =   4590
      End
      Begin VB.Label Label7 
         Caption         =   "Four, Dotted Decimal Bytes"
         Height          =   270
         Index           =   9
         Left            =   2460
         TabIndex        =   27
         Top             =   945
         Width           =   2085
      End
      Begin VB.Label Label7 
         Caption         =   "Unsigned 32 Bit Decimal No."
         Height          =   270
         Index           =   8
         Left            =   2460
         TabIndex        =   26
         Top             =   285
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   $"Form1.frx":0000
         Height          =   480
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   2895
         Width           =   8340
      End
      Begin VB.Label Label1 
         Caption         =   $"Form1.frx":00D8
         Height          =   630
         Left            =   4800
         TabIndex        =   22
         Top             =   360
         Width           =   3540
      End
      Begin VB.Label Label2 
         Caption         =   $"Form1.frx":015F
         Height          =   600
         Left            =   4800
         TabIndex        =   21
         Top             =   1020
         Width           =   3540
      End
      Begin VB.Label Label6 
         Caption         =   "LSByte --------------------------- MSByte   Relative To 32 Bit Dec. No. Above   --- And The Binary No. Just Below."
         Height          =   555
         Left            =   2280
         TabIndex        =   20
         Top             =   1545
         Width           =   2610
      End
      Begin VB.Label Label7 
         Caption         =   "As An Actual Binary Number"
         Height          =   270
         Index           =   0
         Left            =   4995
         TabIndex        =   19
         Top             =   2235
         Width           =   2145
      End
      Begin VB.Label Label8 
         Caption         =   "MSByte           To               LSByte      "
         Height          =   270
         Left            =   1260
         TabIndex        =   18
         Top             =   2565
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Flip High -To- Low Byte Order To Get Actual Internet Address Number Significance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   90
      TabIndex        =   0
      Top             =   3525
      Width           =   8535
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Top             =   450
         Width           =   4590
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   510
         TabIndex        =   3
         Top             =   1215
         Width           =   390
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1515
         TabIndex        =   2
         Top             =   1215
         Width           =   1245
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3255
         TabIndex        =   1
         Top             =   1215
         Width           =   1440
      End
      Begin VB.Label Label7 
         Caption         =   "Class C, Three Hi Bits = 110, 21 Bit Net, 8 Bit Local"
         Height          =   270
         Index           =   12
         Left            =   4785
         TabIndex        =   31
         Top             =   1665
         Width           =   3705
      End
      Begin VB.Label Label7 
         Caption         =   "Class B, Two Hi Bits = 10, 14 Bit Net, 16 Bit Local"
         Height          =   270
         Index           =   11
         Left            =   4800
         TabIndex        =   30
         Top             =   1380
         Width           =   3585
      End
      Begin VB.Label Label7 
         Caption         =   "Class A, Hi Bit = 0, 7 Bit Net, 24 Bit Local Host"
         Height          =   270
         Index           =   10
         Left            =   4800
         TabIndex        =   29
         Top             =   1095
         Width           =   3585
      End
      Begin VB.Label Label3 
         Caption         =   "The Values In This Frame Are Not Used By Windows Programs, But Are The Real Numbers And What They Represent. "
         Height          =   675
         Left            =   4845
         TabIndex        =   28
         Top             =   375
         Width           =   3600
      End
      Begin VB.Label Label7 
         Caption         =   "Binary Bytes"
         Height          =   270
         Index           =   6
         Left            =   1905
         TabIndex        =   11
         Top             =   255
         Width           =   930
      End
      Begin VB.Label Label5 
         Caption         =   "MSByte           To               LSByte"
         Height          =   270
         Left            =   1035
         TabIndex        =   10
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label7 
         Caption         =   "Network Class"
         Height          =   270
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   1575
         Width           =   1200
      End
      Begin VB.Label Label7 
         Caption         =   "Network Number"
         Height          =   225
         Index           =   2
         Left            =   1515
         TabIndex        =   7
         Top             =   1575
         Width           =   1260
      End
      Begin VB.Label Label7 
         Caption         =   "Local Host Number"
         Height          =   270
         Index           =   3
         Left            =   3270
         TabIndex        =   6
         Top             =   1575
         Width           =   1620
      End
      Begin VB.Label Label7 
         Caption         =   "Of /"
         Height          =   225
         Index           =   4
         Left            =   1515
         TabIndex        =   5
         Top             =   1830
         Width           =   1260
      End
      Begin VB.Label Label7 
         Caption         =   "Of /"
         Height          =   225
         Index           =   5
         Left            =   3300
         TabIndex        =   4
         Top             =   1830
         Width           =   1395
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MSDN Library - Visual Basic Documentation - Platform SDK "
      Height          =   255
      Index           =   15
      Left            =   1650
      TabIndex        =   34
      Top             =   6525
      Width           =   4800
   End
   Begin VB.Label Label7 
      Caption         =   "See Note In          General Declarations"
      Height          =   435
      Index           =   14
      Left            =   6750
      TabIndex        =   33
      Top             =   5940
      Width           =   1755
   End
   Begin VB.Label Label7 
      Caption         =   "References:"
      Height          =   270
      Index           =   13
      Left            =   600
      TabIndex        =   32
      Top             =   6030
      Width           =   930
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''      Note:

'''''''      Two Functions In This Example Program Might Be Used In Visual Basic
'''''''      Projects, Instead Of Using Win32 API's. For Address Conversions.


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''' 1.   NetIpLong_ToDot  ...Input A 32 Bit Unsigned No. Address String
'''''''                       ...Returns A Dotted Decimal Address String

'''''''''''' (Win32 API inet_ntoa  Will Also Do This, But It Is A Quirky Function)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''' 2.   DotTo_NetIpLong  ...Input A Dotted Decimal Address String
'''''''                       ...Returns A 32 Bit Unsigned No. Address String

'''''''''''''''       (Win32 API inet_addr  Will Also Do This Well)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''        In Both Functions, You Might Want To Remove The MsgBox Near The Top
'''''''        Of The Function And The -PrntNetSpecs- Sub Call At The Bottom
'''''''        Of Each Function.

Private Nuttin
Public Function NetIpLong_ToDot(ByVal InAddr As String) As String

''Input InAddr, An unsigned Long INet Address String
''Uses Currency Data Type To Hack Out The Four Byte Vals
''Returns Byte.Byte.Byte.Byte "Dotted Decimal" Notation, IP Address String

On Error GoTo BadIn

StarUp1@ = CCur(InAddr)              '''convert to currency

'''4,294,967,263                  '''Largest Posible Valid address

 If StarUp1@ > 4294967263# Then '''''''Added MsgBox For This Prog Only  REM REM
  Mnsg$ = "4,294,967,263 Is The Largest Posible Valid Address"
  MsgBox Mnsg$, 0
 End If ''''''''''''''''''''''''''''''''''''''''''''''''''''''          REM REM

If StarUp1@ < 0 Then                 '''convert to pos. 32 bit unsigned
StarUp1@ = StarUp1@ + 4294967296#
End If

CkekkerCurre@ = CCur(4294967295#)    ''range check val

If StarUp1@ > CkekkerCurre@ Then     ''check range, within 32 bit
StarUp1@ = 0                         ''if bad, set = 0
End If

Rett1@ = CCur(Fix(StarUp1@ / 16777216))                  '''Get High Order Byte

StarUp2@ = CCur(Round(StarUp1@ - (Rett1@ * 16777216)))
Rett2@ = CCur(Fix(StarUp2@ / 65536))                     '''Get Next Byte


StarUp3@ = CCur(Round(StarUp2@ - (Rett2@ * 65536)))
Rett3@ = CCur(Fix(StarUp3@ / 256))                       '''Get Next Byte


Rett4@ = CCur(Round(StarUp3@ - (Rett3@ * 256)))          '''Remainder, Low Order Byte


NetIpLong_ToDot = Trim$(Str$(Rett4@)) & "." & Trim$(Str$(Rett3@)) & "." & Trim$(Str$(Rett2@)) & "." & Trim$(Str$(Rett1@))


GoTo GoodIn
BadIn:
Resume 10
10
NetIpLong_ToDot = "0.0.0.0"          '''If Error, Returns This

GoodIn:

'''''PrntNetSpecs, Below, Added For This Prog Only'''''''''''

PrntNetSpecs Rett1@, Rett2@, Rett3@, Rett4@

End Function
Private Sub Command1_Click()

Text2.Text = NetIpLong_ToDot(Trim$(Text1.Text))

End Sub


Private Sub Command2_Click()

Ding$ = DotTo_NetIpLong(Text2.Text)

Text1.Text = Format$(Ding$, "###,###,###,###,##0")

End Sub



Public Function DotTo_NetIpLong(ByVal InDot As String) As String

''' Input "Dotted Decimal" Notation IP Address String
''' Returns 32 Bit Unsigned Decimal Address As A String Of Digit Characters

Ased1& = InStr(1, InDot, ".", vbTextCompare)

Byt1Lo@ = CCur(Val(Mid$(InDot, 1, Ased1& - 1)))

If Byt1Lo@ > 255 Then
Byt1Lo@ = 255
End If

 If Byt1Lo@ > 223 Then '''''''Added Msg PopFor This Prog Only      REM REM
  Mnsg$ = "223 Is The Largest Posible LSByte In A Valid Address"
  MsgBox Mnsg$, 0
 End If ''''''''''''''''''''''''''''''''''''''''''''''''''''''     REM REM

If Byt1Lo@ < 0 Then
Byt1Lo@ = 0
End If


Ased2& = InStr(Ased1& + 1, InDot, ".", vbTextCompare)

Byt2@ = CCur(Val(Mid$(InDot, Ased1& + 1, (Ased2& - Ased1&) - 1)))

If Byt2@ > 255 Then
Byt2@ = 255
End If

If Byt2@ < 0 Then
Byt2@ = 0
End If

Byt2a@ = Byt2@               '''REM REM this Line Used With PrntNetSpecs Sub
Byt2@ = Byt2@ * 256


Ased3& = InStr(Ased2& + 1, InDot, ".", vbTextCompare)

Byt3@ = CCur(Val(Mid$(InDot, Ased2& + 1, (Ased3& - Ased2&) - 1)))

If Byt3@ > 255 Then
Byt3@ = 255
End If

If Byt3@ < 0 Then
Byt3@ = 0
End If

Byt3a@ = Byt3@               '''REM REM this Line Used With PrntNetSpecs Sub
Byt3@ = Byt3@ * 65536


Byt4Hi@ = CCur(Val(Mid$(InDot, Ased3& + 1)))

If Byt4Hi@ > 255 Then
Byt4Hi@ = 255
End If

If Byt4Hi@ < 0 Then
Byt4Hi@ = 0
End If

Byt4aHi@ = Byt4Hi@               '''REM REM this Line Used With PrntNetSpecs Sub
Byt4Hi@ = Byt4Hi@ * 16777216


AddrssLon@ = Byt1Lo@ + Byt2@ + Byt3@ + Byt4Hi@   ''The Value To Be Returned

DotTo_NetIpLong = Trim$(Str$(AddrssLon@))       ''Converted To A String And Returned


'''''PrntNetSpecs, Below, Added For This Prog Only

PrntNetSpecs Byt4aHi@, Byt3a@, Byt2a@, Byt1Lo@


End Function


Public Function ByteToBinary(ByVal Byytte As String) As String

'''0000 0001

Byt22& = CLng(Byytte)

NxBit8& = (&H80 And Byt22&) / 128

NxBit7& = (&H40 And Byt22&) / 64

NxBit6& = (&H20 And Byt22&) / 32

NxBit5& = (&H10 And Byt22&) / 16

NxBit4& = (&H8 And Byt22&) / 8

NxBit3& = (&H4 And Byt22&) / 4

NxBit2& = (&H2 And Byt22&) / 2

NxBit1& = (&H1 And Byt22&)


ByteToBinary = Trim$(Str$(NxBit8&)) & Trim$(Str$(NxBit7&)) & Trim$(Str$(NxBit6&)) & Trim$(Str$(NxBit5&)) & Trim$(Str$(NxBit4&)) & Trim$(Str$(NxBit3&)) & Trim$(Str$(NxBit2&)) & Trim$(Str$(NxBit1&))

End Function

Public Function BinaryToNumba(ByVal Binsst As String) As String

''in Max 24 Bits Binary String

''ret  Decimal Number String

''0000 0000 0000

StLt% = Len(Binsst)

For Zs% = 1 To StLt%
 Select Case Zs%
 
 Case Is = 1                                  '''LSBit
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno1& = CLng(Sed$)
 Case Is = 2
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno2& = CLng(Sed$) * 2
 Case Is = 3
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno3& = CLng(Sed$) * 4
 Case Is = 4
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno4& = CLng(Sed$) * 8
 Case Is = 5
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno5& = CLng(Sed$) * 16
 Case Is = 6
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno6& = CLng(Sed$) * 32
 Case Is = 7
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno7& = CLng(Sed$) * 64
 Case Is = 8
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno8& = CLng(Sed$) * 128
 Case Is = 9
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno9& = CLng(Sed$) * 256
 Case Is = 10
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno10& = CLng(Sed$) * 512

 Case Is = 11
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno11& = CLng(Sed$) * 1024
 Case Is = 12
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno12& = CLng(Sed$) * 2048
 Case Is = 13
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno13& = CLng(Sed$) * 4096
 Case Is = 14
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno14& = CLng(Sed$) * 8192
 Case Is = 15
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno15& = CLng(Sed$) * 16384
 Case Is = 16
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno16& = CLng(Sed$) * 32768
 Case Is = 17
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno17& = CLng(Sed$) * 65536
 Case Is = 18
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno18& = CLng(Sed$) * 131072
 Case Is = 19
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno19& = CLng(Sed$) * 262144
 Case Is = 20
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno20& = CLng(Sed$) * 524288
 
  Case Is = 21
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno21& = CLng(Sed$) * 1048576
 Case Is = 22
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno22& = CLng(Sed$) * 2097152
 Case Is = 23
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno23& = CLng(Sed$) * 4194304
 Case Is = 24
 Sed$ = Mid$(Binsst, (StLt% - Zs%) + 1, 1)
 Dno24& = CLng(Sed$) * 8388608
 
 End Select

Next Zs%

 tot& = Dno1& + Dno2& + Dno3& + Dno4& + Dno5& + Dno6& + Dno7& + Dno8& + Dno9& + Dno10& + Dno11& + Dno12& + Dno13& + Dno14& + Dno15& + Dno16& + Dno17& + Dno18& + Dno19& + Dno20& + Dno21& + Dno22& + Dno23& + Dno24&
 
 BinaryToNumba = Format$(Trim$(Str$(tot&)), "###,###,###,##0")


End Function


Public Sub PrntNetSpecs(ByVal Rett1@, ByVal Rett2@, ByVal Rett3@, ByVal Rett4@)

'''Class A Network
'' Hi 1 Bit = 0 Binary            Indicates Class A Network
''    7 Bit (0 - 127 Dec)         Network Address
''   24 Bit (0 - 16,777,215 Dec)  Local Host Address

'''Class B Network
''Hi 2 Bits = 10 Binary           Indicates Class B Network
''  14 Bit (0 - 16,383 Dec)       Network Address
''  16 Bit (0 - 65,535 Dec)       Local Host Address

'''Class C Network
''Hi 3 Bits = 110 Binary          Indicates Class B Network
''  21 Bit (0 - 2,097,151 Dec)    Network Address
''   8 Bit (0 - 255 Dec)          Local Host Address

''ex.

'''in Four Curr Bytes From Address in Backward Order

Text3.Text = ByteToBinary(Trim$(Str$(Rett1@))) & " " & ByteToBinary(Trim$(Str$(Rett2@))) & " " & ByteToBinary(Trim$(Str$(Rett3@))) & " " & ByteToBinary(Trim$(Str$(Rett4@)))

'''''Added Network Class Hacked From Top One, Two Three Bits Of LSByte

Df& = InStr(1, ByteToBinary(Trim$(Str$(Rett4@))), "0", vbTextCompare)

ActualBinOrd$ = ByteToBinary(Trim$(Str$(Rett4@))) & ByteToBinary(Trim$(Str$(Rett3@))) & ByteToBinary(Trim$(Str$(Rett2@))) & ByteToBinary(Trim$(Str$(Rett1@)))

Text7.Text = ByteToBinary(Trim$(Str$(Rett4@))) & " " & ByteToBinary(Trim$(Str$(Rett3@))) & " " & ByteToBinary(Trim$(Str$(Rett2@))) & " " & ByteToBinary(Trim$(Str$(Rett1@)))

Select Case Df&
 Case Is = 0
 Text4.Text = "?"
 Text5.Text = "Invalid"
 Text6.Text = "Invalid"
 Label7(4).Caption = "Of / "
 Label7(5).Caption = "Of / "
 Case Is = 1
 Text4.Text = "A"                                         ''Net Class
 Text5.Text = BinaryToNumba(Mid$(ActualBinOrd$, 2, 7))    ''Net No. of 128
 Text6.Text = BinaryToNumba(Mid$(ActualBinOrd$, 9, 24))   ''Host No. 0f 16,777,216
 Label7(4).Caption = "Of / 128"
 Label7(5).Caption = "Of / 16,777,216"
 Case Is = 2
 Text4.Text = "B"                                         ''Net Class
 Text5.Text = BinaryToNumba(Mid$(ActualBinOrd$, 3, 14))   ''Net No. of 16384
 Text6.Text = BinaryToNumba(Mid$(ActualBinOrd$, 17, 16))  ''Host No. of 65536
 Label7(4).Caption = "Of / 16384"
 Label7(5).Caption = "Of / 65536"
 Case Is = 3
 Text4.Text = "C"                                         ''Net Class
 Text5.Text = BinaryToNumba(Mid$(ActualBinOrd$, 4, 21))   ''Net No. of 2,097,152
 Text6.Text = BinaryToNumba(Mid$(ActualBinOrd$, 25, 8))   ''Host No. of 256
 Label7(4).Caption = "Of / 2,097,152"
 Label7(5).Caption = "Of / 256"
 Case Is > 3
 Text4.Text = "?"
 Text5.Text = "Invalid"
 Text6.Text = "Invalid"
 Label7(4).Caption = "Of / "
 Label7(5).Caption = "Of / "
End Select

End Sub

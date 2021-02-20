VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9405
   ClientLeft      =   26055
   ClientTop       =   2025
   ClientWidth     =   19065
   LinkTopic       =   "Form1"
   ScaleHeight     =   627
   ScaleMode       =   3  'ÇÈ¼¿
   ScaleWidth      =   1271
   Begin VB.ListBox List1 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   420
      ItemData        =   "Form1.frx":0000
      Left            =   0
      List            =   "Form1.frx":0002
      TabIndex        =   27
      Top             =   8700
      Width           =   18945
   End
   Begin VB.TextBox txt_Total 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BackColor       =   &H00000000&
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   20.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   525
      Left            =   11970
      TabIndex        =   26
      Text            =   "0"
      Top             =   1410
      Width           =   3045
   End
   Begin VB.CommandButton cmd_Account 
      Caption         =   "Ãâ ±Ý"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   12
      Left            =   14430
      TabIndex        =   25
      Top             =   8070
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ÀÔ ±Ý"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   11
      Left            =   12690
      MaskColor       =   &H00E0E0E0&
      Picture         =   "Form1.frx":0004
      TabIndex        =   24
      Top             =   8070
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      Caption         =   "Ãâ ±Ý"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   10
      Left            =   14430
      TabIndex        =   23
      Top             =   7530
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ÀÔ ±Ý"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   12690
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   22
      Top             =   7530
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      Caption         =   "Ãâ ±Ý"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   14430
      TabIndex        =   21
      Top             =   6990
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ÀÔ ±Ý"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   7
      Left            =   12690
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   20
      Top             =   6990
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      Caption         =   "Ãâ ±Ý"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   14430
      MaskColor       =   &H00808080&
      TabIndex        =   19
      Top             =   6450
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      BackColor       =   &H00FFFF80&
      Caption         =   "ÀÔ ±Ý"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   5
      Left            =   12690
      MaskColor       =   &H00FFFFC0&
      TabIndex        =   18
      Top             =   6450
      UseMaskColor    =   -1  'True
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      Caption         =   "Hopper º¸Ãæ"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   11670
      TabIndex        =   17
      Top             =   4920
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      Caption         =   "Ãâ ±Ý"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   9930
      TabIndex        =   16
      Top             =   4920
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      Caption         =   "Hopper º¸Ãæ"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   11670
      TabIndex        =   15
      Top             =   4380
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      Caption         =   "Ãâ ±Ý"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   9930
      TabIndex        =   14
      Top             =   4380
      Width           =   1485
   End
   Begin VB.CommandButton cmd_Account 
      Caption         =   "ÁöÆó ÀÎÃâ"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   9930
      TabIndex        =   13
      Top             =   3840
      Width           =   1485
   End
   Begin VB.TextBox txt_Bill_H1000_Update 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9900
      TabIndex        =   12
      Text            =   "0"
      Top             =   8070
      Width           =   2000
   End
   Begin VB.TextBox txt_Bill_H5000_Update 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9900
      TabIndex        =   11
      Text            =   "0"
      Top             =   7530
      Width           =   2000
   End
   Begin VB.TextBox txt_Coin_H100_Update 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9900
      TabIndex        =   10
      Text            =   "0"
      Top             =   6990
      Width           =   2000
   End
   Begin VB.TextBox txt_Coin_H500_Update 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9900
      TabIndex        =   9
      Text            =   "0"
      Top             =   6450
      Width           =   2000
   End
   Begin VB.TextBox txt_Bill_H1000 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   6870
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   8070
      Width           =   2000
   End
   Begin VB.TextBox txt_Bill_H5000 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   6870
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   7530
      Width           =   2000
   End
   Begin VB.TextBox txt_Coin_H100 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   6870
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   6990
      Width           =   2000
   End
   Begin VB.TextBox txt_Coin_H500 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   6870
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   6450
      Width           =   2000
   End
   Begin VB.TextBox txt_Coin_C100 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   6870
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   4920
      Width           =   2000
   End
   Begin VB.TextBox txt_Coin_C500 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   6870
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   4380
      Width           =   2000
   End
   Begin VB.TextBox txt_Bill_Stack 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   6870
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   3840
      Width           =   2000
   End
   Begin VB.TextBox txt_Bill_No 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   6870
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   3300
      Width           =   2000
   End
   Begin VB.ComboBox cmb_ParkingName 
      Height          =   300
      Left            =   435
      TabIndex        =   0
      Top             =   690
      Width           =   2565
   End
   Begin Threed.SSCommand cmd_APS_Select 
      Height          =   435
      Left            =   3105
      TabIndex        =   28
      Top             =   645
      Width           =   1290
      _Version        =   65536
      _ExtentX        =   2275
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "Connect"
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "Form1.frx":67F8
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "APS Management"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   465
      Index           =   1
      Left            =   450
      TabIndex        =   66
      Top             =   90
      Width           =   2625
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "'Hopper º¸Ãæ' ¹öÆ°À» ´©¸£°í, °¢ µ¿Àü ¼ö³³ÅëÀÇ µ¿ÀüÀ» ÇØ´ç µ¿Àü¹æÃâ±â¿¡ Ã¤¿ö³Ö½À´Ï´Ù."
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006F3C2F&
      Height          =   645
      Index           =   23
      Left            =   13560
      TabIndex        =   65
      Top             =   4950
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÁöÆóÀÎ½Ä±â ½ºÅÃ¿¡¼­ ÁöÆó¸¦ Ãâ±ÝÇÕ´Ï´Ù...!!"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006F3C2F&
      Height          =   285
      Index           =   22
      Left            =   13560
      TabIndex        =   64
      Top             =   4050
      Width           =   4635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÀÜ°í ÃÑ¾× :"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   20.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Index           =   21
      Left            =   9480
      TabIndex        =   63
      Top             =   1410
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°³"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   19
      Left            =   9000
      TabIndex        =   62
      Top             =   5010
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°³"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   18
      Left            =   9000
      TabIndex        =   61
      Top             =   4470
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "µ¿Àü 100¿ø±Ç"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   17
      Left            =   4650
      TabIndex        =   60
      Top             =   5010
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "µ¿Àü 500¿ø±Ç"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   16
      Left            =   4650
      TabIndex        =   59
      Top             =   4470
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "[ Coin Stacker ] µ¿Àü¼ö³³Åë"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   15
      Left            =   360
      TabIndex        =   58
      Top             =   4380
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "[ Bill Acceptor ] ÁöÆóÀÎ½Ä±â"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   14
      Left            =   360
      TabIndex        =   57
      Top             =   3300
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÁöÆó 1000¿ø±Ç"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   13
      Left            =   4650
      TabIndex        =   56
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÁöÆó 5000¿ø±Ç"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   12
      Left            =   4650
      TabIndex        =   55
      Top             =   7620
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¸Å"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   11
      Left            =   9000
      TabIndex        =   54
      Top             =   8160
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¸Å"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   10
      Left            =   9000
      TabIndex        =   53
      Top             =   7620
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¸Å"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   9
      Left            =   12000
      TabIndex        =   52
      Top             =   8160
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¸Å"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   8
      Left            =   12000
      TabIndex        =   51
      Top             =   7620
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°³"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   7
      Left            =   12000
      TabIndex        =   50
      Top             =   7080
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°³"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   6
      Left            =   12000
      TabIndex        =   49
      Top             =   6540
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°³"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   5
      Left            =   9000
      TabIndex        =   48
      Top             =   7080
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°³"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   4
      Left            =   9000
      TabIndex        =   47
      Top             =   6540
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "µ¿Àü 100¿ø±Ç"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   3
      Left            =   4650
      TabIndex        =   46
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "µ¿Àü 500¿ø±Ç"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   2
      Left            =   4650
      TabIndex        =   45
      Top             =   6540
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "[ Bill Dispenser ] ÁöÆó¹æÃâ±â"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   1
      Left            =   360
      TabIndex        =   44
      Top             =   7560
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "[ Coin Hopper ] µ¿Àü¹æÃâ±â"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Index           =   0
      Left            =   360
      TabIndex        =   43
      Top             =   6480
      Width           =   3795
   End
   Begin VB.Label lbl_Update 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Update Ã³¸®ÀÏ½Ã"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   9540
      TabIndex        =   42
      Top             =   840
      Width           =   6075
   End
   Begin VB.Label lbl_MngName 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "admin"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   405
      Left            =   10710
      TabIndex        =   41
      Top             =   240
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°ü ¸® ÀÚ :"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   405
      Left            =   9510
      TabIndex        =   40
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label lbl_APS 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00000000&
      BorderStyle     =   1  '´ÜÀÏ °íÁ¤
      Caption         =   "±ºÆ÷½Ã¼³°ü¸®°ø´Ü"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   20.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   600
      Left            =   420
      TabIndex        =   39
      Top             =   1290
      Width           =   3960
   End
   Begin VB.Label lbl_Alarm 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00FFFFC0&
      Caption         =   "Áö Æó ÀÎ ½Ä ±â"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   16290
      TabIndex        =   38
      Top             =   120
      Width           =   2475
   End
   Begin VB.Label lbl_Alarm 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00FFFFC0&
      Caption         =   "ÁöÆó¹æÃâ±â 5000¿ø"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   16290
      TabIndex        =   37
      Top             =   810
      Width           =   2475
   End
   Begin VB.Label lbl_Alarm 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00FFFFC0&
      Caption         =   "ÁöÆó¹æÃâ±â 1000¿ø"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   16290
      TabIndex        =   36
      Top             =   1170
      Width           =   2475
   End
   Begin VB.Label lbl_Alarm 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00FFFFC0&
      Caption         =   "µ¿Àü¹æÃâ±â 500¿ø"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   16290
      TabIndex        =   35
      Top             =   1515
      Width           =   2475
   End
   Begin VB.Label lbl_Alarm 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00FFFFC0&
      Caption         =   "µ¿Àü¹æÃâ±â 100¿ø"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   16290
      TabIndex        =   34
      Top             =   1860
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°³"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   20
      Left            =   9000
      TabIndex        =   33
      Top             =   3390
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¿ø"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   24
      Left            =   9000
      TabIndex        =   32
      Top             =   3930
      Width           =   435
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°í°´ ÁöºÒ ±Ý¾× ÀúÀå"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Index           =   4
      Left            =   540
      TabIndex        =   31
      Top             =   2670
      Width           =   2415
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "°Å½º¸§µ· ÁöºÒ ¿¹ºñ±Ý"
      BeginProperty Font 
         Name            =   "¸¼Àº °íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Index           =   5
      Left            =   540
      TabIndex        =   30
      Top             =   5820
      Width           =   2415
   End
   Begin VB.Label lbl_Alarm 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BackColor       =   &H00FFFFC0&
      Caption         =   "µ¿ Àü ¼ö ³³ Åë"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ ExtraBold"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   16290
      TabIndex        =   29
      Top             =   465
      Width           =   2475
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   9150
      TabIndex        =   67
      Top             =   0
      Width           =   9795
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      Height          =   705
      Index           =   0
      Left            =   150
      Shape           =   4  'µÕ±Ù »ç°¢Çü
      Top             =   2490
      Width           =   18675
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
      Height          =   705
      Index           =   1
      Left            =   150
      Shape           =   4  'µÕ±Ù »ç°¢Çü
      Top             =   5640
      Width           =   18675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

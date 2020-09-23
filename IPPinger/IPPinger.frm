VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IP Pinger ßy: Mike Canejo"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4335
   Icon            =   "IPPinger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   -80
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "Ping  "
         Default         =   -1  'True
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Text            =   "24.147.235.241"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   5
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   2880
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "32"
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Loop"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "MIN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Hidden"
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   840
         Width           =   975
      End
      Begin MSWinsockLib.Winsock sock 
         Left            =   3960
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "IP Host:"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes:"
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Quanity:"
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple Pinging Program
'Made ßy: Mike Canejo
'Email: mike_3d@hotmail.com
'
'This is good for keeping IP's or Host addresses alive.
'Or just to quickly check if a website address or IP address exists.





Private Sub Command1_Click()
On Error GoTo here 'If ever, jump to the ending on "HERE:"

Dim x As Integer 'x will be the only variable we need
    
    For x = 1 To Text2 '                    This will open the amount of dos ping windows that you put
                       '                    in the quanity textbox. This is useful to ping a Host address
                       '                    a lot of times.
        
            If Check1.Value = 1 Then '      If the Loop checkbox is checked
            
                If Check2.Value = 1 Then '  If the Hidden checkbox is checked
            
                    Shell "Ping.exe -t -l " & Text3 & " " & Text1, vbHide '         Run the ping.exe looping the ping in hidden mode
                                
                        Else '  if hidden not selected do the below
            
                    Shell "Ping.exe -t -l " & Text3 & " " & Text1, vbNormalFocus '  Run the ping.exe looping theping in normal mode
            
                End If 'End 2nd IF statement
        
            End If ' End 1st IF statement
            
            '*********************************************************************
            '***** TEXT3 holds the amount of bytes to be sent in the pinging *****
            '*********************************************************************
    
            If Check1.Value = 0 Then '      If the Loop checkbox is un-checked
        
                If Check2.Value = 1 Then '  If the Hidden checkbox is checked
        
                    Shell "Ping.exe -l " & Text3 & " " & Text1, vbHide '             Run the ping.exe pinging only once in hidden mode
            
                        Else '  if hidden not selected do the below
        
                    Shell "Ping.exe -l " & Text3 & " " & Text1, vbNormalFocus '      Run the ping.exe pinging only once in normal mode
        
                End If 'End 2nd IF statement
    
            End If ' End 1st IF statement

    Next x ' Return to the For and x loop at the top and do the above
           ' again until it reaches the number specified.

here: 'If any errors, the computer will jump to this line ending the sub

End Sub

Private Sub Command2_Click()

    Text1.SetFocus 'Puts the cursor in text1
    
    Me.WindowState = 1 'Puts the form in the taskbar
    
End Sub

Private Sub Form_Load()

    Text1 = sock.LocalIP 'Just so you can test it out
    
End Sub


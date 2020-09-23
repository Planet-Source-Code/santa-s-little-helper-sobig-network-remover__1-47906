VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSobig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Removal of Sobig.f"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSobig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   13140
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txteIP 
      Height          =   375
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmsStop 
      Caption         =   "&Stop"
      Height          =   855
      Left            =   9240
      Picture         =   "frmSobig.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ListBox lstIPStatus 
      Height          =   2940
      Left            =   8160
      TabIndex        =   11
      Top             =   1800
      Width           =   4815
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   855
      Left            =   11160
      Picture         =   "frmSobig.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox txtEndIP 
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   240
      MaxLength       =   15
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.ListBox lstStatus 
      Height          =   2940
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   7935
   End
   Begin VB.CommandButton cmdCheckMachine 
      Caption         =   "Check Machine"
      Default         =   -1  'True
      Height          =   855
      Left            =   6240
      Picture         =   "frmSobig.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblDot 
      AutoSize        =   -1  'True
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5040
      TabIndex        =   12
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   240
      Left            =   3000
      TabIndex        =   10
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Address Range :-"
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1725
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1155
   End
End
Attribute VB_Name = "frmSobig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject

Dim foundvirus As Boolean
Dim file1 As Boolean
Dim file2 As Boolean
Dim Reg1 As Boolean
Dim Reg2 As Boolean
Dim currentIP As String
Dim sComplete As Boolean

'IP Counters
Dim lTotal As Integer
Dim lFound As Integer
Dim lNot As Integer
Dim lValid As Integer
Dim lWin32 As Integer

Private Sub Final_Message_for_user()
    MsgBox "The subnet has been scanned :-" _
            & vbCrLf & vbCrLf _
            & "Total Address Scanned      : " & lTotal & vcbcrlf & vbCrLf _
            & "Virus Found                        : " & lFound & vbCrLf _
            & "Machines not infected        : " & lNot & vbCrLf _
            & "Invalid I.P Address            : " & lValid & vbCrLf _
            & "Non Win32 Machines          : " & lWin32 & vbCrLf _
            & vbCrLf & "Have a nice day!", vbInformation, "Message"
End Sub

Private Sub cmdCheckMachine_Click()
Dim key1 As String
Dim key2 As String
Dim Registry As clsRegistry
Dim Hoststatus As Boolean
Dim ECHO As ICMP_ECHO_REPLY
Dim pos As Long
Dim success As Long

'Reset Counters
lTotal = "0": lFound = "0": lNot = "0": lValid = "0": lWin32 = "0": pBar.Value = "0"



On Error Resume Next

cmsStop.Enabled = True

Set Registry = New clsRegistry

'Make some checks
If Trim(Len(txtIP)) = 0 Then
    MsgBox "Error - No Starting IP", vbInformation, "Error"
    Exit Sub
End If
If Trim(Len(txtEndIP & "." & txteIP)) = 0 Then
    MsgBox "Error - No Ending IP", vbInformation, "Error"
    Exit Sub
End If

If CheckIP(txtIP) = False Or CheckIP(txtEndIP & "." & txteIP) = False Then
    MsgBox "One of the I.P Address is invalid", vbInformation, "Error"
    Exit Sub
End If

currentIP = txtIP

' Enable the progress bar

pBar.Min = getStartValue(txtIP) ' get the last octet from the starting IP Address
pBar.Max = txteIP

pBar.Visible = True ' Make the progress visible

Me.MousePointer = "11"

While sComplete = False
    ' Current IP Address
    lstStatus.AddItem "IP Address - " & currentIP
    Me.Caption = "Removal of Sobig.f - " & currentIP
    
    If Right$(currentIP, 3) = "254" Then
        sComplete = True
    End If
    
    If currentIP = txtEndIP & "." & txteIP Then
    'Only Exit While when sComplete is equal to True
        sComplete = True
    End If
    
    DoEvents
    
     lstStatus.AddItem "Checking I.P Address - " & currentIP
     lstStatus.Refresh
     
'Checking to see if the IP Responds to a Ping
If SocketsInitialize() Then
   
   'Ping the IP Address and see if you get a response before you continue
   success = Ping((currentIP), "ECHO", ECHO)
   
    ' if the ping works then continue
    If success = "0" Then
        ' If the ping returns a 0 then it the address exists
            SocketsCleanup
            
            If Left$(ECHO.Data, 1) <> Chr$(0) Then
                pos = InStr(ECHO.Data, Chr$(0))
            End If
     ' See if the host will return a name, if it does then it will be a win32 machine
     Hoststatus = GetHostNameFromIP(currentIP) ' I use this function to return the host name of the ip address
     If Hoststatus = True Then
        lblCheck = "Looking for File 'winppr32.exe'"
        lstStatus.AddItem "Looking for file 'winppr32.exe' - " & Now
        If fso.FileExists("\\" & currentIP & "\admin$\winppr32.exe") = True Then
            lblCheck = "Found File 'winppr32.exe'"
            lstStatus.AddItem "======>>>> Found File 'winppr32.exe' - " & Now
            foundvirus = True
            file1 = -True
        End If
        
        lstStatus.Refresh
        
        lblCheck = "Looking for File 'winstt32.dat'"
        lstStatus.AddItem "Looking for file 'winstt32.dat' - " & Now
        
        If fso.FolderExists("\\" & currentIP & "\admin$") = True Then
            If fso.FileExists("\\" & currentIP & "\admin$\winstt32.DAT") = True Then
                lblCheck = "Found File 'winstt32.dat'"
                lstStatus.AddItem "======>>>> Found File 'winstt32.dat' - " & Now
                foundvirus = True
                file2 = True
            End If
            
            lstStatus.Refresh
            
            If file1 = True Or file2 = True Then
                lblCheck = "Checking Regestry HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run 'TrayX' = C:\WINNT\WINPPR32.EXE /sinc"
                lstStatus.AddItem "Checking Regestry HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run 'TrayX' = C:\WINNT\WINPPR32.EXE /sinc" & Now
            
                key1 = Registry.GetValue(eHKEY_LOCAL_MACHINE, "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "TrayX", currentIP)
            
                If Len(key1) <> 0 Then
                    lblCheck = "Found item Regestry HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run 'TrayX' = C:\WINNT\WINPPR32.EXE /sinc"
                    lstStatus.AddItem "======>>>> Found item Regestry HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run 'TrayX' = C:\WINNT\WINPPR32.EXE /sinc" & Now
                    foundvirus = True
                    Reg1 = True
                End If
                
                lstStatus.Refresh
                
                lblCheck = "Checking Regestry HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run 'TrayX' = C:\WINNT\WINPPR32.EXE /sinc"
                lstStatus.AddItem "Checking Regestry HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run 'TrayX' = C:\WINNT\WINPPR32.EXE /sinc" & Now
                
                key2 = Registry.GetValue(eHKEY_CURRENT_USER, "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "TrayX", currentIP)
                
                If Len(key2) <> 0 Then
                    lblCheck = "Found item Regestry HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run 'TrayX' = C:\WINNT\WINPPR32.EXE /sinc"
                    lstStatus.AddItem "======>>>> Found item Regestry HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run 'TrayX' = C:\WINNT\WINPPR32.EXE /sinc" & Now
                    foundvirus = True
                    Reg2 = True
                End If
              End If
            End If
        End If
    End If 'Success
End If 'Socket
        DoEvents
        
    ' if ping status equals 11010 then
    If success = "0" Then
            If Hoststatus = False Then
                lstIPStatus.AddItem currentIP & " - Not a Win32 Machine"
                lWin32 = lWin32 + 1
            Else
                If foundvirus = True Then
                    Remove_Files
                    Remove_Reg Reg1, Reg2
                    lstIPStatus.AddItem currentIP & " - Machine Cleaned"
                    lFound = lFound + 1
                Else
                    lstIPStatus.AddItem currentIP & " - Machine Not Infected"
                    lNot = lNot + 1
                End If
            End If
     Else
           SocketsCleanup
           lstIPStatus.AddItem currentIP & " - Not a Valid I.P Address"
           lValid = lValid + 1
     End If
     lstStatus.Refresh
     lstIPStatus.Refresh
     
    lstStatus.AddItem "*** Finished Checking " & currentIP & " ****"
    
    foundvirus = False: file1 = False: file2 = False: Reg1 = False: Reg2 = False
    
    currentIP = AddOne(currentIP)
    success = "0"

    lstIPStatus.SetFocus
    SendKeys "{End}"
    
    lTotal = lTotal + 1
    pBar.Value = pBar.Value + 1
Wend

Me.MousePointer = vbNormal
lstStatus.AddItem ""
lstStatus.AddItem "*** Finished Checking ****"
lblCheck = "Finished"
sComplete = False

cmsStop.Enabled = False

'Call a message box with all the results
Final_Message_for_user

' Hide the progree bar
pBar.Visible = False

End Sub

Private Sub Remove_Files()
On Error Resume Next
If file1 = True Then
    fso.DeleteFile "\\" & currentIP & "\admin$\winppr32.exe"
    lblCheck = "Removing 'winppr32.exe'"
    lstStatus.AddItem "======>>>> Removing 'winppr32.exe' - " & Now
End If
If file2 = True Then
    fso.DeleteFile "\\" & currentIP & "\admin$\winstt32.DAT"
    lblCheck = "Removing 'winstt32.dat'"
    lstStatus.AddItem "======>>>> Removing 'winstt32.dat' - " & Now
End If

End Sub

Private Function Remove_Reg(Reg1 As Boolean, Reg2 As Boolean)
Dim Regs As clsRegistry

Set Regs = New clsRegistry
If Reg1 = True Then
    Regs.DeleteValue eHKEY_LOCAL_MACHINE, "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "TrayX", currentIP
    lblCheck = "Removing From HKEY_LOCAL_MACHINE"
    lstStatus.AddItem "======>>>> Removing From HKEY_LOCAL_MACHINE - " & Now
End If

If Reg2 = True Then
    Regs.DeleteValue eHKEY_CURRENT_USER, "\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "TrayX", currentIP
    lblCheck = "Removing From HKEY_CURRENT_USER"
    lstStatus.AddItem "======>>>> Removing From HKEY_CURRENT_USER - " & Now
End If

End Function

Private Function AddOne(IP As String) As String
    Dim A As Long, B As Long, C As Long, D As Long
    Dim A1 As Long, B1 As Long, C1 As Long
    A1 = InStr(1, IP, ".")
    A = Mid(IP, 1, A1)
    IP = Mid(IP, A1 + 1, Len(IP) - A1)
    B1 = InStr(1, IP, ".")
    B = Mid(IP, 1, B1)
    IP = Mid(IP, B1 + 1, Len(IP) - B1)
    C1 = InStr(1, IP, ".")
    C = Mid(IP, 1, C1)
    IP = Mid(IP, C1 + 1, Len(IP) - C1)
    D = IP


    If D >= 255 Then


        If C >= 255 Then


            If B >= 255 Then


                If A >= 255 Then
                
                Else
                    A = A + 1
                End If
                B = 0
            Else
                B = B + 1
            End If
            C = 0
        Else
            C = C + 1
        End If
        D = 0
    Else
        D = D + 1
    End If
    AddOne = A & "." & B & "." & C & "." & D
End Function
Private Sub cmdClose_Click()
    Call cmsStop_Click
    End
End Sub

Public Function CheckIP(strIPaddress As String) As Boolean
    Dim strOctet As String
    Dim x As Integer
On Error Resume Next
    For x = 0 To 2
        strOctet = Left(strIPaddress, InStr(1, strIPaddress, ".") - 1)
        strIPaddress = Mid(strIPaddress, InStr(1, strIPaddress, ".") + 1)
        If Not CheckOctet(strOctet) Then
            CheckIP = False
            Exit Function
        End If
    Next
    If Not CheckOctet(strIPaddress) Then
        CheckIP = False
        Exit Function
    End If
    
    CheckIP = True
End Function

Public Function CheckOctet(strOctet As String) As Boolean
    Select Case Len(strOctet)
        Case 1
            If Val(strOctet) < 0 Or Val(strOctet) > 9 Then
                CheckOctet = False
                Exit Function
            End If
        Case 2
            If Val(strOctet) < 10 Or Val(strOctet) > 99 Then
                CheckOctet = False
                Exit Function
            End If
        Case 3
            If Val(strOctet) < 100 Or Val(strOctet) > 255 Then
                CheckOctet = False
                Exit Function
            End If
        Case Else
            CheckOctet = False
            Exit Function
    End Select
    If IsNumeric(strOctet) Then CheckOctet = True
End Function


Private Sub cmsStop_Click()
    sComplete = True
    cmsStop.Enabled = False
End Sub

Private Sub Command1_Click()
   MsgBox getStartValue(txtIP)
End Sub

Private Sub LoadIP()
Dim xlen As Integer
Dim xIP As String
Dim xNew_IP As String
Dim xCurrent As Integer

' Get Length of inital string
xlen = Len(txtIP)

'Get the location of the first .
xCurrent = InStr(txtIP, ".")

'Set the first string


End Sub
Private Function Total_Dots(IP_Address As String) As Integer
Dim n As Integer
Dim xlen As Integer
Dim xtmp As String
Dim dot As Integer
Dim xFinish As Boolean



xtmp = IP_Address: xlen = Len(xtmp): dot = "0"

If Len(xtmp) = 0 Then Exit Function

While xFinish <> True
    
    If Left$(xtmp, 1) = "." Then
        dot = dot + 1
    End If
    
    n = xlen - 1
    
    xtmp = Right$(xtmp, n)
    
    xlen = Len(xtmp)
    
    If xlen < 1 Then
        xFinish = True
    End If
Wend
    Total_Dots = dot

End Function

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Thank you for using my software, if you have any problems plese email me @ " & vbCrLf & vbCrLf _
            & "Danny_rawl@Hotmail.com" & vbCrLf & vbCrLf & "Have a nice day!", vbInformation, "Thank you"
End Sub

Private Sub txteIP_KeyPress(KeyAscii As Integer)

If KeyAscii >= "48" And KeyAscii <= "57" Or KeyAscii = "8" Then

Else
    KeyAscii = "0"
End If


End Sub

Private Sub txtEndIP_KeyPress(KeyAscii As Integer)
If Total_Dots(txtEndIP) <= 2 Then
    
Else
    If KeyAscii <> "8" Then
        KeyAscii = "0"
    End If
End If
End Sub

Private Sub txtIP_Change()

If Total_Dots(txtIP) < 3 Then
    txtEndIP = txtIP
End If
End Sub

Private Sub txtIP_GotFocus()
    txtIP.SelLength = Len(txtIP)
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)



If KeyAscii >= 58 And KeyAscii <= 126 Then
    KeyAscii = "0"
End If

'Between
If Total_Dots(txtIP) <= 3 Then
    If KeyAscii = 46 And Total_Dots(txtIP) = "3" Then
        KeyAscii = "0"
    End If

Else
    If KeyAscii <> "8" Then
        KeyAscii = "0"
    End If
End If
End Sub

Private Function getStartValue(IP_Address As String) As Integer
Dim n As Integer
Dim xlen As Integer
Dim xtmp As String
Dim xFinish As Boolean
Dim xDot As Integer

' This Function will be used to get the last octet of the starting IP Address,
' I am using this so that I can get a starting value for the progress bae

'Reset the Values
xtmp = IP_Address: xlen = Len(xtmp)

' if there is no ip address then there is no point in carrying on
If Len(xtmp) = 0 Then Exit Function

While xFinish <> True
' while there is still a dot in the ip address the while loop is false
    
    
    ' Get the next location of the dot
    xDot = InStr(xtmp, ".")
    
    n = xlen - xDot
    
    xtmp = Right$(xtmp, n)
    
    xlen = Len(xtmp)
    
    If xDot < 1 Then
        xFinish = True
    End If
Wend

'Write the information back to the function
        getStartValue = xtmp
        
End Function

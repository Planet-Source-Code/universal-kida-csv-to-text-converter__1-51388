VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Import_CSV_Form 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CSV to Text..."
   ClientHeight    =   2925
   ClientLeft      =   3420
   ClientTop       =   1965
   ClientWidth     =   7770
   Icon            =   "Import_CSV_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7770
   Begin VB.CommandButton cmdExit4Small 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      Picture         =   "Import_CSV_Form.frx":263A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmd_Browse 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Pick CSV File"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      Picture         =   "Import_CSV_Form.frx":4C74
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   6375
   End
   Begin VB.Frame innerfrm 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   6615
   End
   Begin VB.Frame frmcsvChoose 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select the csv file to Import"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   7575
   End
   Begin ComCtl3.CoolBar bottomBar 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   2505
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   741
      BandCount       =   1
      Picture         =   "Import_CSV_Form.frx":72AE
      BackColor       =   12632256
      _CBWidth        =   7770
      _CBHeight       =   420
      _Version        =   "6.0.8169"
      MinHeight1      =   360
      Width1          =   2880
      NewRow1         =   0   'False
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   741
      BandCount       =   1
      Picture         =   "Import_CSV_Form.frx":752C
      BackColor       =   12632256
      ImageList       =   "ImageList1"
      FixedOrder      =   -1  'True
      _CBWidth        =   7770
      _CBHeight       =   420
      _Version        =   "6.0.8169"
      MinHeight1      =   360
      Width1          =   2880
      NewRow1         =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Import_CSV_Form.frx":77AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Import_CSV_Form.frx":7C43
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   2040
      Width           =   6135
   End
End
Attribute VB_Name = "Import_CSV_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''Purpose : Converting CSV files to Text Files  ''''
''''          as per customisation required       ''''
''''                                              ''''
''''Date : 21st Nov '2003                         ''''
''''                                              ''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''       Date Modified : 27th Dec '2003         ''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmd_Browse_Click()
   ' CancelError is True.
     On Error GoTo errhandler
     Dim txtPath As String
     Dim EnableImport
     cmdImport.Enabled = True

   ' Set filters.
    CommonDialog1.Filter = "All Files (*.*)|*.*|Comma Delimeted Files (*.CSV)|*.CSV|Test Files (*.txt)|*.txt"
   ' Specify default filter.
    CommonDialog1.FilterIndex = 2
    Me.CommonDialog1.InitDir = App.Path
   ' Display the Open dialog box.
    CommonDialog1.ShowOpen
   ' Call the open file procedure.
   'OpenFile (CommonDialog1.FileName)3
    txtPath = Me.CommonDialog1.FileName
txtFileName.Text = txtPath
errhandler:
' User pressed Cancel button.
   Exit Sub

End Sub
Private Sub cmdExit4Small_Click()
Dim Ans As Integer
Ans = MsgBox("Are you sure?", vbYesNo + vbExclamation, "CSV to Text")
If Ans = vbYes Then
' Close the Application
   End
   Unload Me
Else
 ' Cancel
   Exit Sub
End If
End Sub

Private Sub cmdImport_Click()
cmdImport.Enabled = False

If txtFileName.Text = "" Then
    MsgBox "File Not Selected, Please Choose a File to Import"
Else

Dim fso
Dim act
Dim total_imported_text
Set fso = CreateObject("scripting.filesystemobject")
'Set act = fso.OpenTextFile("C:\Documents and Settings\Administrator.BUGS\Desktop\Import CSV\cm11NOV2003bhav.csv")
Set act = fso.OpenTextFile(Me.CommonDialog1.FileName)
total_imported_text = act.ReadAll
total_imported_text = Replace(total_imported_text, Chr(13), "*")
total_imported_text = Replace(total_imported_text, Chr(10), "*")
'Response.Write total_imported_text
total_imported_text = Replace(total_imported_text, Chr(34), "")
'Remove all the quotes (If your csv has quotes other than to seperate text
'You may want to remove this modifier to the imported text
total_split_text = Split(total_imported_text, "*")
'Split the file up by comma
total_num_imported = UBound(total_split_text)
For i = 1 To total_num_imported - 1 '0 To total_num_imported '
    comma_split = Split(total_split_text(i), ",")
      On Error Resume Next
    If comma_split(0) <> "" Then
        Fileld2OfExcel = Trim(Mid(comma_split(0), 2))
        '****************Existing Condition*******************************
        'Check the column of the excel sheets if it is empty
        'if not then print then Row
        '****************As per Your Condition*******************************
        'A new text file will be created for each row with the text file name as the first column
        If Fileld2OfExcel <> "" Then
           '****************************************************
           'Debug.Print total_split_text(i)
           '****************************************************
           'Save Each Next Row that is Found
           If Dir(App.Path & "\Data\" & comma_split(0) & ".txt") = "" Then
               Open App.Path & "\Data\" & comma_split(0) & ".txt" For Output As #1
                   Print #1, Format$(comma_split(10), "dd/mm/yyyy") & "," & comma_split(2) & "," & comma_split(3) & "," & comma_split(4) & "," & comma_split(5) & "," & comma_split(8) 'Mid(total_split_text(i), 2) & vbCrLf
               Close #1
           Else
                L_B_Found = False
                Less_D_Found = False
                Dim myData As String
                Dim dt1 As Date
                Dim dt2 As Date
                myData = ""
                myData1 = ""
                Open App.Path & "\Data\" & comma_split(0) & ".txt" For Input As #1
                Do While Not EOF(1)
                    Line Input #1, myData
                    If myData <> "" Then
                       txt = Split(myData, ",")
                       dt1 = CDate(comma_split(10))
                       dt2 = CDate(Format(txt(0), "dd/mm/yyyy"))
                       If dt1 = dt2 Then
                          L_B_Found = True
                          myData1 = myData1 & myData & vbCrLf
                       ElseIf dt1 < dt2 Then
                          If Less_D_Found = False Then
                             Less_D_Found = True
                             myData1 = myData1 & Format$(comma_split(10), "dd/mm/yyyy") & "," & comma_split(2) & "," & comma_split(3) & "," & comma_split(4) & "," & comma_split(5) & "," & comma_split(8) & vbCrLf & myData & vbCrLf
                          Else
                          L_B_Found = False
                             myData1 = myData1 & myData & vbCrLf
                          End If
                       Else
                          myData1 = myData1 & myData & vbCrLf & Format$(comma_split(10), "dd/mm/yyyy") & "," & comma_split(2) & "," & comma_split(3) & "," & comma_split(4) & "," & comma_split(5) & "," & comma_split(8) & vbCrLf
                       End If
                    End If
                Loop
                Close #1
                If L_B_Found = False Then
                    Open App.Path & "\Data\" & comma_split(0) & ".txt" For Output As #1
                         Print #1, myData1
                    Close #1
              End If
           End If
        End If
     End If
Next i

Label1.Caption = "Recently Converted File : " & Me.CommonDialog1.FileTitle
MsgBox Me.CommonDialog1.FileTitle & " Has Been Imported"
End If
End Sub

Private Sub Form_Load()
Label1.Caption = "No File Converted in this Session, Choose 'Pick CSV File' Button To Convert"
cmdImport.Enabled = False
     
DName = App.Path & "\Data"

Dim sDummy As String
On Error Resume Next
    CMK = False
If Right(DName, 1) <> "\" Then
DName = DName & "\"
sDummy = Dir$(DName & "*.*", vbDirectory)
DirExists = Not (sDummy = "")
End If

If DirExists = "True" Then
    CMK = False
Else
CMK = True
MkDir App.Path & "\Data"
End If
End Sub

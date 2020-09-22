VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ini Example"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Ini Filepath"
      Height          =   735
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   3345
      Begin VB.TextBox txtFilePath 
         Height          =   330
         Left            =   90
         TabIndex        =   9
         Top             =   270
         Width           =   3075
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Other"
      Height          =   2400
      Left            =   90
      TabIndex        =   7
      Top             =   2475
      Width           =   4425
      Begin VB.CommandButton Command9 
         Caption         =   "Make It Pretty"
         Height          =   375
         Left            =   180
         TabIndex        =   24
         Top             =   1935
         Width           =   1725
      End
      Begin VB.CommandButton Command8 
         Caption         =   "All Section Names"
         Height          =   375
         Left            =   180
         TabIndex        =   23
         Top             =   1530
         Width           =   1725
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Single String"
         Height          =   375
         Left            =   180
         TabIndex        =   22
         Top             =   1125
         Width           =   1725
      End
      Begin VB.ListBox List1 
         Height          =   1815
         Left            =   2295
         TabIndex        =   21
         Top             =   315
         Width           =   1860
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Collection"
         Height          =   375
         Left            =   180
         TabIndex        =   20
         Top             =   720
         Width           =   1725
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Array"
         Height          =   375
         Left            =   180
         TabIndex        =   19
         Top             =   315
         Width           =   1725
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "String"
      Height          =   735
      Left            =   90
      TabIndex        =   5
      Top             =   900
      Width           =   4470
      Begin VB.CommandButton Command1 
         Caption         =   "load"
         Height          =   285
         Left            =   3510
         TabIndex        =   14
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Caption         =   "save"
         Height          =   285
         Left            =   2835
         TabIndex        =   15
         Top             =   315
         Width           =   645
      End
      Begin VB.TextBox txtName 
         Height          =   330
         Left            =   1755
         TabIndex        =   12
         Text            =   "Lewis Miller"
         Top             =   315
         Width           =   915
      End
      Begin VB.TextBox txtNameKey 
         Height          =   330
         Left            =   495
         TabIndex        =   11
         Text            =   "Name"
         Top             =   315
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Value:"
         Height          =   195
         Left            =   1260
         TabIndex        =   13
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "Key:"
         Height          =   285
         Left            =   90
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Number"
      Height          =   735
      Left            =   90
      TabIndex        =   1
      Top             =   1710
      Width           =   4425
      Begin VB.TextBox txtAgeKey 
         Height          =   330
         Left            =   495
         TabIndex        =   17
         Text            =   "Age"
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         Caption         =   "load"
         Height          =   285
         Left            =   3510
         TabIndex        =   2
         Top             =   270
         Width           =   645
      End
      Begin VB.TextBox txtAge 
         Height          =   330
         Left            =   1755
         TabIndex        =   4
         Text            =   "31"
         Top             =   270
         Width           =   915
      End
      Begin VB.CommandButton Command4 
         Caption         =   "save"
         Height          =   285
         Left            =   2835
         TabIndex        =   3
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Value:"
         Height          =   195
         Left            =   1260
         TabIndex        =   18
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label3 
         Caption         =   "Key:"
         Height          =   285
         Left            =   90
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Section"
      Height          =   735
      Left            =   3510
      TabIndex        =   0
      Top             =   90
      Width           =   1095
      Begin VB.TextBox txtSection 
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Text            =   "Settings"
         Top             =   270
         Width           =   825
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Please note: See Command7 for new comments about WriteSection()

'this is all you need to do to use the ini
'class file in your program,
'(and make sure you add the class file to your program)
Dim Ini As New cINI

Private Sub Command1_Click()

  'retrieve a string value from the ini file

  'section, key, and default value if not found

    txtName = Ini.ReadString(txtSection.Text, txtNameKey.Text, "Lewis Miller")

End Sub

Private Sub Command2_Click()

  'save the name to an ini file
  'in this case txtname, but it can be
  'a label,string,caption etc

    Ini.WriteString txtSection.Text, txtNameKey.Text, txtName.Text

End Sub

Private Sub Command3_Click()

  Dim lngMyAge As Long

    'load a number from an ini file
    lngMyAge = Ini.ReadNumber(txtSection.Text, txtAgeKey.Text, 31)

    'convert the number to a string so
    'we can display it in a text box
    txtAge.Text = CStr(lngMyAge)

End Sub

Private Sub Command4_Click()

  Dim lngMyAge As Long

    'save a number to an ini file
    'in this example we are reading
    'the number from a textbox,so we
    'have to convert it to a number (from a string)
    'first check to make sure its a number
    If Len(txtAge.Text) > 0 Then
        If IsNumeric(txtAge.Text) Then
            lngMyAge = CLng(txtAge.Text)
        End If
    End If

    Ini.WriteNumber txtSection.Text, txtAgeKey.Text, lngMyAge

End Sub

Private Sub Command5_Click()

  Dim strArr() As String
  Dim X As Long

    'shows how easy it is to save and load an array
    'with this ini class

    'fill the array with junk for
    ' this example
    ReDim strArr(8)
    For X = 0 To 8
        strArr(X) = String$(20, X + 65)
    Next X

    'save the array
    Ini.WriteArray "MyArray", strArr

    'erase the array and the list box
    Erase strArr
    List1.Clear

    'now load the array from the ini file again
    ' this will = true if there was any saved items
    If Ini.ReadArray("MyArray", strArr) = True Then
        For X = 0 To UBound(strArr)
            List1.AddItem strArr(X)
        Next X
    End If

End Sub

Private Sub Command6_Click()

  Dim colNew As Collection
  Dim X As Long

    'shows how to save and load collections
    ' similar to an array, except collections
    'arent passed byref

    'fill collection with junk for this example
    Set colNew = New Collection
    For X = 0 To 6
        colNew.Add String$(20, X + 65)
    Next X

    'save the collection
    Ini.WriteCollection "MyCollection", colNew

    'erase the collection and the listbox
    Set colNew = New Collection
    List1.Clear

    'load the collection from the ini file
    Set colNew = Ini.ReadCollection("MyCollection")

    'show the collection in the listbox
    If colNew.Count > 0 Then
        For X = 1 To colNew.Count
            List1.AddItem colNew(X)
        Next X
    End If

    'now we can delete an entire section
    'note: it doesnt delete the section name only the sections sub settings
    '      you can use DeleteSection() to delete an entire section
    Ini.WriteSection "MyCollection", ""

End Sub

Private Sub Command7_Click()

  'shows how to save a simple string
  ' to an ini file

  Dim strMyString As String

    strMyString = InputBox("Enter a String.", "", "This is my cool string")

    'save the string
    'this is also handy for deleting entire sections at once
    'by using an empty string for the value (see command6)
    'first delete the section (see note below)
    Ini.DeleteSection "MyString"
    'save it
    Ini.WriteSection "MyString", strMyString

    'erase the string
    strMyString = ""

    'reload it from the ini file
    strMyString = Ini.ReadSection("MyString")
    
    MsgBox strMyString

'Note:
'when you write a string to a section that already contains strings
'with WriteSection() your string is popped to the first in the list.
'When you retrieve that section again with ReadSection() each string
'each string will be seperated by a Chr$(0) (or vbnullchar) starting from
'last saved to the first string saved
'ex: if you did this-
    'Ini.WriteSection "MyString", "This is my cool string"
    'Ini.WriteSection "MyString", "This is another cool string"

'then if you read the section like this -
   '    strMyString = Ini.ReadSection("MyString")

' strMyString would look like this:
'  This is another cool string vbNullChar This is my cool string vbNullChar

'in the ini file it looks like this (from last saved to first saved)
'[MyString]
'This is another cool string
'This is my cool string

'you can keep it from doing this by first deleting the section before saving a string
'or use it to your advantage to save string lists

End Sub

Private Sub Command8_Click()

  Dim strSectionNames As String
  Dim lngPlace As Long

    'read all section names from the ini file
    strSectionNames = Ini.ReadSectionNames("", txtFilePath.Text) '<< this can be empty
    'or same result
    strSectionNames = Ini.ReadSectionNames '<< see? :)

    'this call can be used in conjunction with
    'writesection to delete everything in the ini file
    'or find all the section names in an unknown ini file

    'all the sectionnames are returned as a string with
    ' each name followed by a chr$(0), so we can split it into
    ' a string array or whatever, since vb5 doesnt have split() we
    ' will just show it with a loop

    List1.Clear

    While Len(strSectionNames) > 0
        lngPlace = InStr(strSectionNames, Chr$(0))
        List1.AddItem Left$(strSectionNames, lngPlace - 1)
        strSectionNames = Mid$(strSectionNames, lngPlace + 1)
    Wend

End Sub

Private Sub Command9_Click()

  'this function will put spaces between each section in your ini
  'file making it more human readable, you can specify an ini file
  'in the parameter argument, or leave it blank for current ini file

    Ini.MakeIniPretty Ini.FilePath
    
    'look at it
    Call Shell("notepad.exe " & Ini.FilePath, vbNormalFocus)

End Sub

Private Sub Form_Load()

  ' the ini class file provides a default path
  ' so lets use it, you can change it to whatever
  ' you wish to use if the default isnt right for you

    txtFilePath.Text = Ini.FilePath

End Sub


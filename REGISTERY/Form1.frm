VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   495
   ClientTop       =   915
   ClientWidth     =   6255
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6255
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5160
      Top             =   2160
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Special Folders"
      TabPicture(0)   =   "Form1.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option1(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option1(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option1(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Option1(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Option1(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Option1(9)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Option1(10)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Option1(11)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Option1(12)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Option1(13)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Option1(14)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Option1(15)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Option1(16)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Option1(17)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Option1(18)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Option1(19)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "the_path"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Previous Search"
      TabPicture(1)   =   "Form1.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "reset_search"
      Tab(1).Control(1)=   "alternative"
      Tab(1).Control(2)=   "Previous_search"
      Tab(1).Control(3)=   "Replace_search"
      Tab(1).Control(4)=   "delete_search"
      Tab(1).Control(5)=   "Label4"
      Tab(1).Control(6)=   "Label2"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Url visited"
      TabPicture(2)   =   "Form1.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "new_entry"
      Tab(2).Control(1)=   "Replace_url"
      Tab(2).Control(2)=   "Command2"
      Tab(2).Control(3)=   "open_web"
      Tab(2).Control(4)=   "delete_url"
      Tab(2).Control(5)=   "typed_urls"
      Tab(2).Control(6)=   "Label5"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "HomePage"
      TabPicture(3)   =   "Form1.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label6"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label7"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label8"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Command8"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Command7"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "homepage"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "openhomepage"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      Begin VB.CommandButton openhomepage 
         Caption         =   "Open The Home Page"
         Height          =   615
         Left            =   1440
         TabIndex        =   39
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox new_entry 
         Height          =   615
         Left            =   -73680
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   38
         Top             =   4560
         Width           =   3135
      End
      Begin VB.CommandButton Replace_url 
         Caption         =   "Replace This Url  By :"
         Height          =   615
         Left            =   -73680
         TabIndex        =   37
         Top             =   3840
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete all Url's (Reset History)"
         Height          =   615
         Left            =   -73680
         TabIndex        =   36
         Top             =   3120
         Width           =   3135
      End
      Begin VB.TextBox homepage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   35
         Top             =   1680
         Width           =   4455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "CHANGE"
         Height          =   375
         Left            =   1560
         TabIndex        =   34
         ToolTipText     =   "CHANGE THE URL IN THE ABOVE TEXT BOX TO THE DESIRED URL AND THEN PRESS CHANGE TO SET YOUR NEW HOMEPAGE"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "USE BLANK"
         Height          =   375
         Left            =   3480
         TabIndex        =   33
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton open_web 
         Caption         =   "Open The Web Page"
         Height          =   495
         Left            =   -73680
         TabIndex        =   31
         Top             =   1920
         Width           =   3135
      End
      Begin VB.CommandButton reset_search 
         Caption         =   "Delete all the above Records ( Reset  The  History )"
         Height          =   735
         Left            =   -74280
         TabIndex        =   30
         ToolTipText     =   "Reset The History of The keywords you searched with before"
         Top             =   3480
         Width           =   4095
      End
      Begin VB.TextBox alternative 
         Height          =   375
         Left            =   -70920
         TabIndex        =   28
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox the_path 
         Height          =   615
         Left            =   -74520
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   27
         Top             =   4080
         Width           =   4815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sendto"
         Height          =   375
         Index           =   19
         Left            =   -71280
         TabIndex        =   26
         Top             =   2100
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Start Menu"
         Height          =   375
         Index           =   18
         Left            =   -71280
         TabIndex        =   25
         Top             =   2580
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Startup"
         Height          =   375
         Index           =   17
         Left            =   -71280
         TabIndex        =   24
         Top             =   3060
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Templates"
         Height          =   375
         Index           =   16
         Left            =   -71280
         TabIndex        =   23
         Top             =   3540
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "history"
         Height          =   375
         Index           =   15
         Left            =   -74400
         TabIndex        =   22
         Top             =   3060
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Programs"
         Height          =   375
         Index           =   14
         Left            =   -71280
         TabIndex        =   21
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "My Music"
         Height          =   375
         Index           =   13
         Left            =   -72720
         TabIndex        =   20
         Top             =   1620
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "desktop"
         Height          =   375
         Index           =   12
         Left            =   -72720
         TabIndex        =   19
         Top             =   2340
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nethood"
         Height          =   375
         Index           =   11
         Left            =   -72720
         TabIndex        =   18
         Top             =   2700
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Personal"
         Height          =   375
         Index           =   10
         Left            =   -72720
         TabIndex        =   17
         Top             =   3060
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Printhood"
         Height          =   375
         Index           =   9
         Left            =   -72720
         TabIndex        =   16
         Top             =   3540
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Local AppData"
         Height          =   375
         Index           =   8
         Left            =   -74400
         TabIndex        =   15
         Top             =   3540
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "fonts"
         Height          =   375
         Index           =   7
         Left            =   -74400
         TabIndex        =   14
         Top             =   2700
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Administrative Tools"
         Height          =   375
         Index           =   6
         Left            =   -74400
         TabIndex        =   13
         Top             =   1260
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "desktop"
         Height          =   375
         Index           =   5
         Left            =   -74400
         TabIndex        =   12
         Top             =   1980
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Local Settings"
         Height          =   375
         Index           =   4
         Left            =   -72720
         TabIndex        =   11
         Top             =   1260
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Recent"
         Height          =   375
         Index           =   3
         Left            =   -71280
         TabIndex        =   10
         Top             =   1620
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "My Pictures"
         Height          =   375
         Index           =   2
         Left            =   -72720
         TabIndex        =   9
         Top             =   1980
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "favorites"
         Height          =   375
         Index           =   1
         Left            =   -74400
         TabIndex        =   8
         Top             =   2340
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "AppData"
         Height          =   375
         Index           =   0
         Left            =   -74400
         TabIndex        =   7
         Top             =   1620
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Open The Folder"
         Height          =   615
         Left            =   -73320
         TabIndex        =   6
         Top             =   4800
         Width           =   1935
      End
      Begin VB.CommandButton delete_url 
         Caption         =   "Delete Url From Record"
         Height          =   495
         Left            =   -73680
         TabIndex        =   5
         Top             =   2520
         Width           =   3135
      End
      Begin VB.ComboBox typed_urls 
         Height          =   315
         Left            =   -74400
         TabIndex        =   4
         Text            =   "Typed Urls"
         Top             =   1440
         Width           =   4695
      End
      Begin VB.ComboBox Previous_search 
         Height          =   315
         Left            =   -74400
         TabIndex        =   3
         Text            =   "Previous Searched Keywords"
         Top             =   1260
         Width           =   4815
      End
      Begin VB.CommandButton Replace_search 
         Caption         =   "Replace Record"
         Height          =   495
         Left            =   -71760
         MaskColor       =   &H000080FF&
         TabIndex        =   2
         ToolTipText     =   $"Form1.frx":037A
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CommandButton delete_search 
         Caption         =   "Delete Record"
         Height          =   495
         Left            =   -74400
         TabIndex        =   1
         ToolTipText     =   "To delete an Item from the list above , choose the Item from the List then click ""Delete Record"""
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Vote Here......"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         MouseIcon       =   "Form1.frx":040D
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "WAITING FOR YOUR COMMENTS AND VOTING"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   44
         Top             =   4200
         Width           =   5415
      End
      Begin VB.Label Label6 
         Caption         =   $"Form1.frx":0717
         Height          =   615
         Left            =   360
         TabIndex        =   43
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label Label5 
         Caption         =   $"Form1.frx":07C9
         Height          =   735
         Left            =   -74760
         TabIndex        =   42
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label4 
         Caption         =   $"Form1.frx":08A4
         Height          =   735
         Left            =   -74520
         TabIndex        =   41
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   $"Form1.frx":0964
         Height          =   735
         Left            =   -74640
         TabIndex        =   40
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "HomePage:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   32
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "with:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71520
         TabIndex        =   29
         Top             =   2520
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Namee As String
Dim url As String
Dim START_PAGE As String
Dim Search_path As String
Dim m As Integer
Dim get_path As String
Dim string_ext As String
Dim i As Integer
Dim url_vote As String
Dim counter As Integer
Dim scrooling As String
Dim text_to_scrool As String
Dim computer As String
Dim COUNTER1 As Integer




Private Sub Command2_Click()
    On Error Resume Next
    For m = 0 To 24
        Call RegdelKey(url & m)
        Call PreviousUrls
    Next m
    Call MsgBox("All Url's are Deleted , 'History Reset' ", vbInformation, "I Love Registery")
End Sub



Private Sub Command6_Click()
    Call RunBrowser(the_path.Text, 10, 1)
End Sub

Private Sub Command7_Click()
    Call RegCreateKey(START_PAGE, homepage.Text)
    Call Home_page
    Call MsgBox("Done !! , Your New Homepage is Set to " & homepage.Text, vbInformation, "I Love Registery")
End Sub

Private Sub Command8_Click()
    Call RegCreateKey(START_PAGE, "ABOUT:BLANK")
    Call Home_page
    Call MsgBox("Done !! , Your New Homepage is Set to " & homepage.Text, vbInformation, "I Love Registery")
End Sub



Private Sub delete_search_Click()
    Call path_search_item(Previous_search.ListIndex, Namee)
    Call RegdelKey(Namee)
    Call PreviousSearch
    Call MsgBox("Item Deleted", vbInformation, "I Love Registery")
End Sub

Private Sub delete_url_Click()
    Call RegdelKey(url & typed_urls.ListIndex + 1)
    Call PreviousUrls
    Call MsgBox("Url Deleted", vbInformation, "I Love Registery")
End Sub

Private Sub Form_Load()
    START_PAGE = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Start Page"
    url = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\TypedURLs\url"
    Search_path = "HKEY_CURRENT_USER\Software\Microsoft\Search Assistant\ACMru\560"
    Call get_computer_name(computer)
    Call PreviousSearch
    Call PreviousUrls
    Call Home_page
End Sub
Private Sub RegCreateKey(Folder As String, value As String)
    Dim OBJ As Object
    On Error Resume Next
    Set OBJ = CreateObject("wscript.shell")
    OBJ.regwrite Folder, value
End Sub

Private Function RegreadKey(value As String) As String
    Dim OBJ As Object, r As String
    r = ""
    On Error GoTo exitt:
    Set OBJ = CreateObject("wscript.shell")
    RegreadKey = OBJ.regread(value)
exitt:
End Function

Private Function RegdelKey(value As String) As String
    Dim OBJ As Object, r As String
    r = ""
    On Error GoTo exitt:
    Set OBJ = CreateObject("wscript.shell")
    r = OBJ.Regdelete(value)
exitt:
End Function

Private Sub Label8_Click()
url_vote = "http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=42785&lngWId=1"
Call RunBrowser(url_vote, 10, 1)
End Sub

Private Sub open_web_Click()
    Call RunBrowser(typed_urls.List(typed_urls.ListIndex), 10, 1)
End Sub

Private Sub openhomepage_Click()
    Call RunBrowser(homepage.Text, 10, 1)
End Sub

Private Sub Option1_Click(Index As Integer)

    get_path = RegreadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" & Option1(Index).Caption)

    the_path.Text = get_path

End Sub

Function Home_page()
    homepage.Text = RegreadKey(START_PAGE)
End Function


Private Sub PreviousSearch()
    Previous_search.Clear
    For m = 0 To 48
        If m < 10 Then string_ext = "3\00" & m
        If m > 9 And m < 25 Then string_ext = "3\0" & m
        If m > 24 And m < 24 + 10 Then string_ext = "4\00" & m - 24
        If m > 33 And m < 49 Then string_ext = "4\0" & m - 24
    
         Namee = Search_path & string_ext
 
        Call Previous_search.AddItem(RegreadKey(Namee), m)
    Next m
    Previous_search.Text = "Previous Searched Keywords"
End Sub


Private Sub PreviousUrls()
    typed_urls.Clear
    For m = 1 To 25
        Call typed_urls.AddItem(RegreadKey(url & m), m - 1)
    Next m
    typed_urls.Text = "Typed Urls"
End Sub

Function path_search_item(m, Namee)
    If m < 10 Then string_ext = "3\00" & m
    If m > 9 And m < 25 Then string_ext = "3\0" & m
    If m > 24 And m < 24 + 10 Then string_ext = "4\00" & m - 24
    If m > 33 And m < 49 Then string_ext = "4\0" & m - 24
    Namee = Search_path & string_ext
End Function

Private Sub Replace_search_Click()
    Call path_search_item(Previous_search.ListIndex, Namee)
    Call RegCreateKey(Namee, alternative.Text)
    Call PreviousSearch
    Call MsgBox("Item Replaced By " & alternative.Text, vbInformation, "I Love Registery")
    alternative.Text = ""
End Sub

Private Sub Replace_url_Click()
    Call RegCreateKey(url & typed_urls.ListIndex + 1, new_entry.Text)
    Call PreviousUrls
    Call MsgBox("Url Replaced Successfully", vbInformation, "I Love Registery")
End Sub

Private Sub reset_search_Click()
    For i = 1 To Previous_search.ListCount
        Call path_search_item(i, Namee)
        Call RegdelKey(Namee)
        Call PreviousSearch
        Call MsgBox("All Items Deleted !! , Search List Reset", vbInformation, "I Love Registery")
    Next i
End Sub




Private Sub Timer1_Timer()
counter = counter + 1
If counter > Len(text_to_scrool) Then
counter = 0
End If
scrooling = Right(text_to_scrool, Len(text_to_scrool) - counter)
Form1.Caption = scrooling
End Sub

Function get_computer_name(computer)
Dim compname As String * 256
Call GetComputerName(compname, 256)
computer = Left(compname, InStr(compname, Chr(0)) - 1)
text_to_scrool = "Hi, " & computer & " & Welcome to the Program ,  This program is a collection of the capabilities of using the Registery To Get Data about Your PC and Writing Data Again, It Conatins The Following : 1) Special Folder's Path's for ex:( 'Desktop' , 'Programs' , 'Favorites' , 'History' , ......etc ) , 2)The Previous Keywords Used in Search , This words can't be removed , Now You can Remove it Very Easy or even replace it , 3) The History of URL Visited , You can remove only only URL without Loosing the Rest , 4) Changing the HomePage , ALL THAT WITHOUT ANY API'S  , A FREE API CODE Designed by 'Sherif Rofael' in 29-th jan. 2003 mailto:ya3amo@hotmail.com , Thanks " & computer & " for using the program , BYE............."
For COUNTER1 = 1 To 30
text_to_scrool = Chr(32) & text_to_scrool
Next COUNTER1
End Function

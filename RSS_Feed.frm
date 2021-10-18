VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmTesteRss 
   Caption         =   "Leitor de RSS"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   13380
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstEndereco 
      Height          =   645
      ItemData        =   "RSS_Feed.frx":0000
      Left            =   120
      List            =   "RSS_Feed.frx":0037
      TabIndex        =   6
      Top             =   240
      Width           =   3855
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   4935
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   13095
      ExtentX         =   23098
      ExtentY         =   8705
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Height          =   975
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "RSS_Feed.frx":02DA
      Top             =   2400
      Width           =   9135
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "RSS_Feed.frx":02E0
      Top             =   1080
      Width           =   9135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "RSS_Feed.frx":02E6
      Top             =   600
      Width           =   9135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "RSS_Feed.frx":02EC
      Top             =   240
      Width           =   9135
   End
   Begin VB.Label Label1 
      Caption         =   "RSS Feeds"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmTesteRss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rssXML As Object

Private Sub CarregaFeed()

    Dim Item As Object
    Dim i As Long
    
    List1.Clear
    
    Set rssXML = CreateObject("msxml2.domdocument")
    rssXML.async = False
    rssXML.Load (lstEndereco.Text)
    Set Item = rssXML.getElementsByTagName("item")

    For i = 0 To Item.length - 1
        On Error Resume Next
        List1.AddItem Item(i).getElementsByTagName("title").Item(0).firstChild.nodeValue
    Next i

End Sub

Private Sub List1_Click()
    
    On Error Resume Next
    Text1 = rssXML.getElementsByTagName("item").Item(List1.ListIndex).getElementsByTagName("title").Item(0).firstChild.nodeValue
    Text2 = rssXML.getElementsByTagName("item").Item(List1.ListIndex).getElementsByTagName("link").Item(0).firstChild.nodeValue
    Text3 = rssXML.getElementsByTagName("item").Item(List1.ListIndex).getElementsByTagName("description").Item(0).firstChild.nodeValue
    Text4 = rssXML.getElementsByTagName("item").Item(List1.ListIndex).xml

    WebBrowser.Silent = True
    WebBrowser.Navigate "about:blank"
    DoEvents
    WebBrowser.Document.Write "<html>" _
                             & "    <head>" _
                             & "        <h3>" & Text1 & "</h3>" _
                             & "    </head>" _
                             & "     <body>" & "" _
                             & "         " & Replace(Text4, "]]>", "") & "" _
                             & "    </body>" _
                             & "</html>"

End Sub

Private Sub List1_Scroll()

    List1_Click

End Sub

Private Sub lstEndereco_Click()
    
    Screen.MousePointer = 11
    
    CarregaFeed
    
    Screen.MousePointer = 0

End Sub

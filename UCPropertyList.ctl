VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl UCPropertyList 
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   ScaleHeight     =   7395
   ScaleWidth      =   5025
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnCol 
      Caption         =   "..."
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cboEdit 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown-Liste
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TxtEdit 
      Appearance      =   0  '2D
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ListView LVPData 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView LVObjData 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   1931
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   3087
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'Kein
      DrawStyle       =   5  'Transparent
      Height          =   255
      Left            =   3360
      ScaleHeight     =   255
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "UCPropertyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_lvSelItem As ListItem     '//the selected LV-Item
Dim m_LVIsSetting As Boolean    '//Flag to avoid raising the event

'//Where do we find our values orparameters
Private Const LVDATACOL_TAG = 0
Private Const LVDATACOL_TEXT = 1
Private Const LVDATACOL_VALUE = 2
Private Const LVDATACOL_CTLTYPE = 3
Private Const LVDATACOL_DATATYPE = 4
Private Const LVDATACOL_OBJINDEX = 5 '//Used for Using 1 Combo for different Indexes
Private Const LVDATACOL_DATALOCKED = 6

'//Our two columns that show Data
Private Const COLHEADER_TEXT = 2
Private Const COLHEADER_VALUE = 3

'//This is for the selection
Private Const LVOBJECTITEMDATA_INDEX = 0
Private Const LVOBJECTITEMDATA_VALUE = 1
Private Const LVOBJECTITEMDATA_TEXT = 2

'//The only Event we need
'//Just to say that something has changed
'//AND
'//WHAT
Public Event ValueChanged(TAG As String, vValue As Variant)

Private m_WidthColShow As Long
Private m_WidthColSpace As Long

Public Enum ShowType
    ShowTypeText = 1
    ShowTypeCombo = 2

    ShowTypeColor = 3

End Enum

Public Enum DataType
    DataTypeText = 1
    DataTypeNumber = 2
    DataTYpeBool = 3

    DataTypeColor = 4

End Enum

Private m_bHasColors As Boolean
'
''//so we need
''//     the Items to be shown (left side of the Listview)
''//     The Corresponding Value/Item of the Object
''//     Type to be shown (text/Combo/... ?)
''//     The Data to be put to the Combo : Text and Numeric Value
'
'
'//Here the Text for the showlist is given .... and some parameters
Public Function AddItem(ItemTAG As String, vShowType As ShowType, Showtext As String, ValueType As DataType, Optional OptionIndexd As Long = 0, Optional NoEdit As Boolean = False) As Long
Dim lvitem As ListItem
    If ItemTAG = "0" Then
         Error (0)
    End If
    Set lvitem = LVPData.ListItems.Add(, , ItemTAG)
    lvitem.SubItems(LVDATACOL_TEXT) = Showtext
    lvitem.SubItems(LVDATACOL_CTLTYPE) = vShowType
    lvitem.SubItems(LVDATACOL_DATATYPE) = ValueType
    lvitem.SubItems(LVDATACOL_DATALOCKED) = -Int(NoEdit)
    lvitem.SubItems(LVDATACOL_OBJINDEX) = OptionIndexd
    lvitem.TAG = ItemTAG
    AddItem = lvitem.Index
End Function
'
'
'//Here our values are set to showlist
'//Additional we can set a LockValue maybe we need to lock upon values
Public Sub SetItemData(ItemTAG As String, ItemValue As Variant, Optional NoEdit As Boolean = False)
Dim lvitem As ListItem
Dim ivb As Long '//Temporary Var for value
    Set lvitem = LVPData.FindItem(ItemTAG)
    TxtEdit.Visible = False
    cboEdit.Visible = False
    btnCol.Visible = False
    If Not lvitem Is Nothing Then

        Select Case lvitem.SubItems(LVDATACOL_CTLTYPE)
            Case ShowTypeText
                lvitem.SubItems(LVDATACOL_VALUE) = ItemValue
                lvitem.SubItems(LVDATACOL_DATALOCKED) = -NoEdit
            Case ShowTypeCombo
                lvitem.SubItems(LVDATACOL_DATALOCKED) = -NoEdit
                '//Find the value to be shown
                If lvitem.SubItems(LVDATACOL_DATATYPE) = DataTYpeBool Then
                    ivb = -ItemValue
                Else
                    ivb = ItemValue
                End If
                If Val(lvitem.SubItems(LVDATACOL_OBJINDEX)) Then
                    '//We took the first way and initialized with predefined values
                    lvitem.SubItems(LVDATACOL_VALUE) = FindComboShowVal((lvitem.SubItems(LVDATACOL_OBJINDEX)), ivb)
                Else
                    '//or we took it from INdex (second way)
                    lvitem.SubItems(LVDATACOL_VALUE) = FindComboShowVal(lvitem.Index, ivb)
                End If
                '//Store the selection value
                '//... so we have two things to store (Text info and a value)
                '//We put the numeric value in the Tag-property of the column
                lvitem.ListSubItems(LVDATACOL_VALUE).TAG = ivb

                '//something nice :
                '//If we have a bool value we colorize by value
                '//note in form2 we had two different calls to this
                '//1. Yes =1
                '//2. Yes =0
                '//so if the value (=1) represents true we give green
                '//And ... on 1.) yes is green (true) , on 2.) no is green(true)
                If lvitem.SubItems(LVDATACOL_DATATYPE) = DataTYpeBool Then
                    If Val(lvitem.ListSubItems(LVDATACOL_VALUE).TAG) = -True Then
                        lvitem.ListSubItems(LVDATACOL_VALUE).ForeColor = vbGreen
                    Else
                        lvitem.ListSubItems(LVDATACOL_VALUE).ForeColor = vbRed
                    End If
                End If
            Case ShowTypeColor
                If m_bHasColors = False Then
                    m_bHasColors = True
                End If
                lvitem.SubItems(LVDATACOL_DATALOCKED) = -NoEdit
                lvitem.SubItems(LVDATACOL_VALUE) = Space(m_WidthColSpace) & "H&" & Hex(ItemValue)
                lvitem.ListSubItems(LVDATACOL_VALUE).TAG = ItemValue
                'lvitem.ListSubItems(LVDATACOL_VALUE).ForeColor = ItemValue
                SetPicLines
        End Select

    End If
    If m_bHasColors Then
        Picture1.Visible = True


        LVPData.Picture = Picture1.Image
    Else
        Picture1.Visible = False
        Set LVPData.Picture = Nothing

    End If
End Sub

Private Sub SetPicLines()
Dim BarHeight  As Long  '/* height of 1 line in the listview
Dim BarWidth   As Long  '/* width of listview
Dim diff        As Long  '/* used in calculations of row height
Dim twipsy      As Long  '/* variable holding Screen.TwipsPerPicture1elY
Dim n As Long
twipsy = Screen.TwipsPerPixelY
Picture1.Top = LVPData.ListItems(1).Top
Picture1.Left = LVPData.Left + LVPData.ColumnHeaders(COLHEADER_VALUE).Left
Picture1.Width = LVPData.ColumnHeaders(COLHEADER_VALUE).Width
Picture1.Height = LVPData.Height
BarHeight = LVPData.ListItems(1).Height
BarWidth = LVPData.ColumnHeaders(COLHEADER_VALUE).Width
    LVPData.PictureAlignment = lvwTopRight
    Picture1.BackColor = vbWhite
    For n = 1 To LVPData.ListItems.Count
        If LVPData.ListItems(n).SubItems(LVDATACOL_CTLTYPE) = ShowTypeColor Then
            Picture1.Line (-LVPData.ColumnHeaders(COLHEADER_VALUE).Left, (n - 1) * BarHeight)-(m_WidthColShow, (n - 1) * BarHeight + BarHeight), Val(LVPData.ListItems(n).ListSubItems(LVDATACOL_VALUE).TAG), BF
        Else
            Picture1.Line (LVPData.ColumnHeaders(COLHEADER_VALUE).Left, (n - 1) * BarHeight)-(LVPData.ColumnHeaders(COLHEADER_VALUE).Left + BarWidth, (n - 1) * BarHeight + BarHeight), vbWhite, BF
        End If
    Next
    Picture1.AutoSize = True
    Picture1.Refresh
    LVPData.Picture = Picture1.Image
End Sub
'//Add an selection Item
Public Sub AddSelectionData(SelectionIndex As Long, ObjectValue As Variant, Showtext As String)
Dim lvitem As ListItem
    Set lvitem = LVObjData.ListItems.Add(, , CStr(SelectionIndex))
    lvitem.SubItems(LVOBJECTITEMDATA_VALUE) = ObjectValue
    lvitem.SubItems(LVOBJECTITEMDATA_TEXT) = Showtext

End Sub

'//Clear the data to be shown
Public Sub ClearData()
Dim n As Long
    TxtEdit.Visible = False
    cboEdit.Visible = False
    For n = 1 To LVPData.ListItems.Count
        LVPData.ListItems(n).SubItems(LVDATACOL_VALUE) = ""
    Next
End Sub

'//Clear a Group of Selection data...
'//Maybe for reloading values
Public Sub ClearSelectionType(SelectionID As Long)
Dim lvitem As ListItem, n As Long
    Do
        Set lvitem = LVObjData.FindItem(SelectionID, lvwText)
        If Not lvitem Is Nothing Then
            n = lvitem.Index
            LVObjData.ListItems.Remove n
        End If
    Loop While Not lvitem Is Nothing
End Sub



'//search in selection values for corresponding data
'//LVPData.FindItem doesn't work because we have two items to check
Private Function FindComboShowVal(ItemIndex As Long, ItemValue As Variant) As String
Dim n As Long
    For n = 1 To LVObjData.ListItems.Count
        Debug.Print LVObjData.ListItems(n), LVObjData.ListItems(n).SubItems(1), LVObjData.ListItems(n).SubItems(2)
        If Val(LVObjData.ListItems(n)) = ItemIndex And LVObjData.ListItems(n).SubItems(LVOBJECTITEMDATA_VALUE) = ItemValue Then
            FindComboShowVal = LVObjData.ListItems(n).SubItems(LVOBJECTITEMDATA_TEXT)
            Exit For
        End If
    Next n
End Function
'
'//set a combobox to a given value
Public Sub SetCboToItemData(pCbo As ComboBox, vItem As Long)
Dim n As Long
    For n = 0 To pCbo.ListCount - 1
        If pCbo.ItemData(n) = vItem Then '//We stored this index
            pCbo.ListIndex = n
            Exit For
        End If
    Next n
End Sub
'
Private Sub btnCol_Click()
Dim n As Long
    CommonDialog1.ShowColor
    n = CommonDialog1.Color
    
    'm_lvSelItem.SubItems(LVDATACOL_VALUE)
    If n >= 0 Then

'        TxtEdit.BackColor = n
        Debug.Print Hex$(n)
        m_lvSelItem.SubItems(LVDATACOL_VALUE) = Space(m_WidthColSpace) & "H&" & Hex(n)
        m_lvSelItem.ListSubItems(LVDATACOL_VALUE).TAG = n
        m_LVIsSetting = True
        TxtEdit.Text = "H&" & Hex(n)
        m_LVIsSetting = False
        DoEvents
        SetPicLines
        DoEvents
        RaiseEvent ValueChanged(m_lvSelItem.TAG, Val(m_lvSelItem.ListSubItems(LVDATACOL_VALUE).TAG))
    End If
End Sub

Private Sub cboEdit_Click()
    m_lvSelItem.SubItems(LVDATACOL_VALUE) = cboEdit.Text
    m_lvSelItem.ListSubItems(LVDATACOL_VALUE).TAG = cboEdit.ItemData(cboEdit.ListIndex)
    If m_lvSelItem.SubItems(LVDATACOL_DATATYPE) = DataTYpeBool Then
        If Val(m_lvSelItem.ListSubItems(LVDATACOL_VALUE).TAG) = -True Then
            m_lvSelItem.ListSubItems(LVDATACOL_VALUE).ForeColor = vbGreen
        Else
            m_lvSelItem.ListSubItems(LVDATACOL_VALUE).ForeColor = vbRed
        End If
        If m_LVIsSetting Then Exit Sub
        RaiseEvent ValueChanged(m_lvSelItem.TAG, CBool(Val(m_lvSelItem.ListSubItems(LVDATACOL_VALUE).TAG)))
        
    Else
        If m_LVIsSetting Then Exit Sub
        RaiseEvent ValueChanged(m_lvSelItem.TAG, Val(m_lvSelItem.ListSubItems(LVDATACOL_VALUE).TAG))
    
    End If
End Sub

Private Sub LVPData_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Set m_lvSelItem = Item
    TxtEdit.Visible = False
    cboEdit.Visible = False
    btnCol.Visible = False
    LVPData.SelectedItem.Selected = False
    Set LVPData.SelectedItem = Nothing
    If Val(Item.SubItems(LVDATACOL_DATALOCKED)) Then Exit Sub
    
    Select Case Val(Item.SubItems(LVDATACOL_CTLTYPE))
        Case ShowTypeText
            
            m_LVIsSetting = True    '//Set FLAG
            TxtEdit.Text = Item.SubItems(LVDATACOL_VALUE)
            
            
            '//position control
            TxtEdit.Left = LVPData.Left + LVPData.ColumnHeaders(COLHEADER_VALUE).Left + 40
            TxtEdit.Width = LVPData.ColumnHeaders(COLHEADER_VALUE).Width - 30
            TxtEdit.Top = LVPData.Top + (Item.Index - 1) * Item.Height + 20
            TxtEdit.Height = Item.Height
            TxtEdit.BackColor = vbWhite
            TxtEdit.Visible = True
            TxtEdit.SetFocus
            m_LVIsSetting = False '//Re-Set FLAG
        Case ShowTypeCombo
            m_LVIsSetting = True
            '//... again :
            '//in case of different selection-types (predefined or by index)
            If Val(Item.SubItems(LVDATACOL_OBJINDEX)) Then
                FillCombo Val(Item.SubItems(LVDATACOL_OBJINDEX))
            Else
                FillCombo Item.Index
            End If
            '//and give the Combo it's show value
            '//... so we have two things to store (Text info and a value)
            '//We put the numeric value in the Tag-property of the column
            SetCboToItemData cboEdit, Val(Item.ListSubItems(LVDATACOL_VALUE).TAG)
            
            '//position control
            cboEdit.Left = LVPData.Left + LVPData.ColumnHeaders(COLHEADER_VALUE).Left + 20
            cboEdit.Width = LVPData.ColumnHeaders(COLHEADER_VALUE).Width - 10
            cboEdit.Top = LVPData.Top + (Item.Index - 1) * Item.Height + 20
'            cboEdit.Height = Item.Height
            cboEdit.Visible = True
            cboEdit.SetFocus
            m_LVIsSetting = False
        Case ShowTypeColor
            m_LVIsSetting = True    '//Set FLAG
            TxtEdit.Text = Trim(Item.SubItems(LVDATACOL_VALUE))
            
            
            '//position control
            TxtEdit.Top = LVPData.Top + (Item.Index - 1) * Item.Height + 40
            btnCol.Top = TxtEdit.Top
            
            
            
            TxtEdit.Left = LVPData.Left + LVPData.ColumnHeaders(COLHEADER_VALUE).Left + 50 + m_WidthColShow
            TxtEdit.Width = LVPData.ColumnHeaders(COLHEADER_VALUE).Width - 30 - btnCol.Width - m_WidthColShow
            btnCol.Left = TxtEdit.Left + TxtEdit.Width
            
            TxtEdit.Top = LVPData.Top + (Item.Index - 1) * Item.Height + 40
            TxtEdit.Height = Item.Height
            'TxtEdit.BackColor = Val(Item.ListSubItems(LVDATACOL_VALUE).TAG)
            TxtEdit.Visible = True
            btnCol.Visible = True
            TxtEdit.SetFocus
            
            m_LVIsSetting = False '//Re-Set FLAG
            
    End Select
End Sub

'//Fill the vlaues to be shown in the combo
Private Sub FillCombo(ItemIndex)
Dim n As Long
    cboEdit.Clear
    For n = 1 To LVObjData.ListItems.Count
        'Debug.Print n, LVObjData.ListItems(n), LVObjData.ListItems(n).SubItems(1), LVObjData.ListItems(n).SubItems(2)
        If Val(LVObjData.ListItems(n)) = ItemIndex Then
            With cboEdit
                .AddItem LVObjData.ListItems(n).SubItems(2)
                '//Give it the index that we expect to return
                .ItemData(.NewIndex) = LVObjData.ListItems(n).SubItems(1)
            End With
        End If
    Next
End Sub

'//Just notify that no Selection is made
Private Sub LVPData_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set m_lvSelItem = LVPData.HitTest(x, y)
    If m_lvSelItem Is Nothing Then
        TxtEdit.Visible = False
        cboEdit.Visible = False
        btnCol.Visible = False
        If Not LVPData.SelectedItem Is Nothing Then
            LVPData.SelectedItem.Selected = False
        End If
        Set LVPData.SelectedItem = Nothing
        'ClearData
    End If
End Sub

'//tie textbox and listviewitem together
Private Sub TxtEdit_Change()
    If m_LVIsSetting Then Exit Sub
    m_lvSelItem.SubItems(LVDATACOL_VALUE) = TxtEdit.Text
    '//if we are just initializing all values then we don't need to raise the event
    '//occurs on selecting a listview item
    
    If m_lvSelItem.SubItems(LVDATACOL_DATATYPE) = DataTypeColor Then
        RaiseEvent ValueChanged(m_lvSelItem.TAG, Val(m_lvSelItem.ListSubItems(LVDATACOL_VALUE).TAG))
    Else
        RaiseEvent ValueChanged(m_lvSelItem.TAG, m_lvSelItem.SubItems(LVDATACOL_VALUE))
    End If
End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0 '//Supress Beep on Enter
        TxtEdit.Visible = False
        LVPData.SetFocus
        '//And Report our Form that something was changed
        RaiseEvent ValueChanged(m_lvSelItem.TAG, m_lvSelItem.SubItems(LVDATACOL_VALUE))
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        KeyAscii = 0 '//Supress Beep on Enter
        TxtEdit.Visible = False
        LVPData.SetFocus
    End If
End Sub

Private Sub UserControl_Initialize()
Dim s0 As String
    m_WidthColShow = Picture1.TextWidth("WW")   '//Show width for colorbar
    
    '//And here we calculate the amount of spaces to be added to the Listview column
    s0 = ""
    m_WidthColSpace = 0
    Do
        s0 = s0 & " "
        m_WidthColSpace = m_WidthColSpace + 1
        Debug.Print Picture1.TextWidth(s0)
    Loop While Picture1.TextWidth(s0) < m_WidthColShow
    
End Sub

Private Sub UserControl_Resize()
Dim n As Integer
'//Just set the Listview to usercontrols size
    LVPData.Top = 0
    LVPData.Left = 0
    LVPData.Width = UserControl.ScaleWidth
    LVPData.Height = UserControl.ScaleHeight
    '//and set the columns widths
    For n = 1 To LVPData.ColumnHeaders.Count
        If n = COLHEADER_TEXT Then
            LVPData.ColumnHeaders(n).Width = UserControl.ScaleWidth / 2 - 25
        ElseIf n = COLHEADER_VALUE Then
            LVPData.ColumnHeaders(n).Width = UserControl.ScaleWidth / 2 - 25
        Else
            LVPData.ColumnHeaders(n).Width = 0
        End If
    Next
End Sub





'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Gibt eine Zugriffsnummer (von Microsoft Windows) für den Gerätekontext des Objekts zurück."
    hDC = UserControl.hDC
End Property


'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Gibt eine Zugriffsnummer (von Microsoft Windows) auf ein Objektfenster zurück."
    hWnd = UserControl.hWnd
End Property




Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub


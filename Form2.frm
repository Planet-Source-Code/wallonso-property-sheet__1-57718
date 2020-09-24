VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form2"
   ScaleHeight     =   4965
   ScaleWidth      =   7680
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton btnResult 
      Caption         =   "See Results Class 1"
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton btnSel1 
      Caption         =   "assume a selection Class2"
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtResult 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton btnResult 
      Caption         =   "See Results Class 1"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton btnSel1 
      Caption         =   "assume a selection Class1"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin PropPage.UCPropertyList UCPList 
      Height          =   1575
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   3615
      _extentx        =   6376
      _extenty        =   2778
   End
   Begin PropPage.UCPropertyList UCPList2 
      Height          =   1575
      Left            =   3480
      TabIndex        =   5
      Top             =   2520
      Width           =   3615
      _extentx        =   6376
      _extenty        =   2778
   End
   Begin VB.Label Label3 
      Caption         =   "Long Style"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Short Style"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "change something"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//Index for Selection-Types
Private Const PROPSEL_YESNO = 1
Private Const PROPSEL_NOYES = 2
Private Const PROPSEL_RGB = 3

'//Index for Selection Value
Private Const SEL_YN_YES = 1
Private Const SEL_YN_NO = 0

'//Index for showlines
Private Const PROPPERTY1 = 1
Private Const PROPPERTY2 = 2
Private Const PROPPERTY3 = 3
Private Const PROPPERTY4 = 4
Private Const PROPPERTY5 = 5
Private Const PROPPERTY6 = 6

'//The Vars we will work with
Dim tclass1 As New ClassTest
Dim tclass2 As New ClassTest
Dim tActC As New ClassTest


'//Just show what we have in our Class
'//Whatever you will do later with it
Private Sub btnResult_Click(Index As Integer)
    txtResult = ""
    Select Case Index
        Case 0
            txtResult.Text = txtResult & tclass1.LockedVal
            txtResult.Text = txtResult & vbCrLf & tclass1.ProT2
            txtResult.Text = txtResult & vbCrLf & tclass1.SunIsShining
            txtResult.Text = txtResult & vbCrLf & tclass1.ColorVal
            txtResult.Text = txtResult & vbCrLf & tclass1.ThisWasFun
        Case 1
            txtResult.Text = txtResult & tclass2.LockedVal
            txtResult.Text = txtResult & vbCrLf & tclass2.ProT2
            txtResult.Text = txtResult & vbCrLf & tclass2.SunIsShining
            txtResult.Text = txtResult & vbCrLf & tclass2.ColorVal
            txtResult.Text = txtResult & vbCrLf & tclass2.ThisWasFun
    End Select
End Sub

'//Set our actual Class Item and show und PropList
Private Sub btnSel1_Click(Index As Integer)
    Select Case Index
        Case 0
            Set tActC = tclass1
            ClassToUCPL UCPList, tActC
            ClassToUCPL UCPList2, tActC
        Case 1
            Set tActC = tclass2
            ClassToUCPL UCPList, tActC
            ClassToUCPL UCPList2, tActC
    End Select
    
End Sub

'//Set the Vlues in the PropList
Sub ClassToUCPL(pUCPL As UCPropertyList, tcls As ClassTest)
    pUCPL.SetItemData PROPPERTY1, tcls.LockedVal
    pUCPL.SetItemData PROPPERTY2, tcls.ProT2
    pUCPL.SetItemData PROPPERTY3, -tcls.SunIsShining
    pUCPL.SetItemData PROPPERTY4, tcls.ColorVal
    pUCPL.SetItemData PROPPERTY5, -tcls.ThisWasFun
End Sub

Private Sub Command1_Click()
    UCPList.ClearSelectionType PROPSEL_RGB
End Sub

Private Sub Form_Load()
Dim res As Long
'//This is One Style we can Use
'//Here we first define our combo-values or slection values
'//I prefer this
    UCPList.AddSelectionData PROPSEL_YESNO, SEL_YN_YES, "Yes"
    UCPList.AddSelectionData PROPSEL_YESNO, SEL_YN_NO, "No"

    UCPList.AddSelectionData PROPSEL_NOYES, SEL_YN_YES, "No"
    UCPList.AddSelectionData PROPSEL_NOYES, SEL_YN_NO, "Yes"

    UCPList.AddSelectionData PROPSEL_RGB, vbRed, "RED"
    UCPList.AddSelectionData PROPSEL_RGB, vbGreen, "Green"
    UCPList.AddSelectionData PROPSEL_RGB, vbBlue, "Blue"
'//The the Lines to be shown
    UCPList.AddItem PROPPERTY1, ShowTypeText, "I'm Locked", DataTypeText, , True
    UCPList.AddItem PROPPERTY2, ShowTypeText, "Type some Text", DataTypeText, , False
    UCPList.AddItem PROPPERTY3, ShowTypeCombo, "Is it nice", DataTYpeBool, PROPSEL_YESNO, False
    UCPList.AddItem PROPPERTY4, ShowTypeColor, "The Color is", DataTypeColor, PROPSEL_RGB, False
    UCPList.AddItem PROPPERTY5, ShowTypeCombo, "This is Fun", DataTYpeBool, PROPSEL_NOYES, False


'//And This the other way
    '//Showline
    UCPList2.AddItem PROPPERTY1, ShowTypeText, "I'm Locked", DataTypeText, , True
    UCPList2.AddItem PROPPERTY2, ShowTypeText, "Type some Text", DataTypeText, , False
    '//Get LineINdex
    res = UCPList2.AddItem(PROPPERTY3, ShowTypeCombo, "Is it nice", DataTYpeBool, , False)
    '//And use it as Index for selection value
        UCPList2.AddSelectionData res, SEL_YN_YES, "Yes"
        UCPList2.AddSelectionData res, SEL_YN_NO, "No"
    
    res = UCPList2.AddItem(PROPPERTY4, ShowTypeColor, "The Color is", DataTypeColor, , False)
'''        UCPList2.AddSelectionData res, vbRed, "RED"
'''        UCPList2.AddSelectionData res, vbGreen, "Green"
'''        UCPList2.AddSelectionData res, vbBlue, "Blue"
    
    res = UCPList2.AddItem(PROPPERTY5, ShowTypeCombo, "This is fun", DataTYpeBool, , False)
        UCPList2.AddSelectionData res, SEL_YN_NO, "Yes"
        UCPList2.AddSelectionData res, SEL_YN_YES, "No"
    
    
    
'//Now init the Classes and give them any values
    Set tclass1 = New ClassTest
    Set tclass2 = New ClassTest
    tclass1.LockedVal = 10
    tclass2.LockedVal = 12
    tclass1.Prop1 = "1"
    tclass2.Prop1 = "2"
    tclass1.ProT2 = "(1)Prop1"
    tclass2.ProT2 = "(2)Prop1"
    tclass1.SunIsShining = False
    tclass2.SunIsShining = True
    tclass1.ColorVal = vbGreen
    tclass2.ColorVal = vbRed
    tclass1.ThisWasFun = True
    tclass2.ThisWasFun = True
End Sub

'//Her our first list indicates thta a value has changed
Private Sub UCPList_ValueChanged(TAG As String, vValue As Variant)
If tActC Is Nothing Then Exit Sub '//Just to make sure tha no error comes
'//The TAG is the index of the show-lines we gave them on init
    Select Case Val(TAG)
        Case PROPPERTY1
            tActC.LockedVal = vValue
            '//Just for Test : say the other PropList the new value
            UCPList2.SetItemData PROPPERTY1, tActC.LockedVal
        Case PROPPERTY2
            tActC.ProT2 = vValue
            UCPList2.SetItemData PROPPERTY2, tActC.ProT2
        Case PROPPERTY3
            tActC.SunIsShining = CBool(vValue)
            UCPList2.SetItemData PROPPERTY3, tActC.SunIsShining
        Case PROPPERTY4
            tActC.ColorVal = vValue
            UCPList2.SetItemData PROPPERTY4, tActC.ColorVal
        Case PROPPERTY5
            tActC.ThisWasFun = CBool(vValue)
            UCPList2.SetItemData PROPPERTY5, tActC.ThisWasFun
    End Select
End Sub

Private Sub UCPList2_ValueChanged(TAG As String, vValue As Variant)
If tActC Is Nothing Then Exit Sub
    Select Case Val(TAG)
        Case PROPPERTY1
            tActC.LockedVal = vValue
            UCPList.SetItemData PROPPERTY1, tActC.LockedVal
        Case PROPPERTY2
            tActC.ProT2 = vValue
            UCPList.SetItemData PROPPERTY2, tActC.ProT2
        Case PROPPERTY3
            tActC.SunIsShining = CBool(vValue)
            UCPList.SetItemData PROPPERTY3, tActC.SunIsShining
        Case PROPPERTY4
            tActC.ColorVal = vValue
            UCPList.SetItemData PROPPERTY4, tActC.ColorVal
        Case PROPPERTY5
            tActC.ThisWasFun = CBool(vValue)
            UCPList.SetItemData PROPPERTY5, tActC.ThisWasFun
    End Select

End Sub

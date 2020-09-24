VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'//Just to show something

Private mvarProp1 As Integer
Private mvarColorVal As Long
Private mvarSunIsShining As Boolean
Private mvarThisWasFun As Boolean
Private mvarProT2 As String
Private mvarLockedVal As Long
Public Property Let LockedVal(ByVal vData As Long)
    mvarLockedVal = vData
End Property
Public Property Get LockedVal() As Long
    LockedVal = mvarLockedVal
End Property



Public Property Let ProT2(ByVal vData As String)
    mvarProT2 = vData
End Property
Public Property Get ProT2() As String
    ProT2 = mvarProT2
End Property



Public Property Let ThisWasFun(ByVal vData As Boolean)
    mvarThisWasFun = vData
End Property
Public Property Get ThisWasFun() As Boolean
    ThisWasFun = mvarThisWasFun
End Property



Public Property Let SunIsShining(ByVal vData As Boolean)
    mvarSunIsShining = vData
End Property
Public Property Get SunIsShining() As Boolean
    SunIsShining = mvarSunIsShining
End Property



Friend Property Let ColorVal(ByVal vData As Long)
    mvarColorVal = vData
End Property
Friend Property Get ColorVal() As Long
    ColorVal = mvarColorVal
End Property



Public Property Let Prop1(ByVal vData As Integer)
    mvarProp1 = vData
End Property
Public Property Get Prop1() As Integer
    Prop1 = mvarProp1
End Property




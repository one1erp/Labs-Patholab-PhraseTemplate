VERSION 5.00
Begin VB.UserControl PhraseTemplateCtrl 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   ScaleHeight     =   345
   ScaleWidth      =   2295
   Begin VB.ComboBox cmbPhraseTemplate 
      Height          =   360
      ItemData        =   "PhraseTemplateCtrl.ctx":0000
      Left            =   0
      List            =   "PhraseTemplateCtrl.ctx":0002
      TabIndex        =   0
      Text            =   "cmbPhraseTemplate"
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "PhraseTemplateCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Connection As ADODB.Connection
Private rst As ADODB.Recordset
Public InitContent As String
Public PhraseName As String
Public Event OnChange()
Private DictPhrase As New Dictionary
'search combobox
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
Private Const CB_ERR = -1, CB_SELECTSTRING = &H14D, CB_SHOWDROPDOWN = &H14F, CBN_SELENDOK = 9

Private Sub TxtFreeTextTemplate_Change()
    RaiseEvent OnChange
End Sub

Private Sub TxtFreeTextTemplate_KeyPress(KeyAscii As Integer)
End Sub

Private Sub cmbPhraseTemplate_click()
    SetBackColor
End Sub

Private Sub cmbPhraseTemplate_DropDown()
    cmbPhraseTemplate.BackColor = vbWhite
    cmbPhraseTemplate.Refresh
End Sub
'call this function in KeyPress event method
Private Function AutoMatchCBBox(ByRef cbBox As ComboBox, ByVal KeyAscii As Integer) As Integer
    
        
    Dim strFindThis As String, bContinueSearch As Boolean, bContinueSearchAfterDelete As Boolean
    Dim originalText As String
    
    Dim lResult As Long, lStart As Long, lLength As Long
    AutoMatchCBBox = 0 ' block cbBox since we handle everything
    bContinueSearch = True
    lStart = cbBox.SelStart
    lLength = cbBox.SelLength

    On Error GoTo ErrHandle
    originalText = ""
        
    If KeyAscii < 32 Then 'control char
        bContinueSearch = False
        cbBox.SelLength = 0 'select nothing since we will delete/enter
        If KeyAscii = Asc(vbBack) Then  'take care BackSpace and Delete first
           
           
            If lLength = 0 Then 'delete last char
                If Len(cbBox) > 0 Then ' in case user delete empty cbBox
                    cbBox.Text = Left(cbBox.Text, Len(cbBox) - 1)
                 
                End If
            Else 'leave unselected char(s) and delete rest of text
                cbBox.Text = Left(cbBox.Text, lStart)
                
            End If
                bContinueSearch = True
                bContinueSearchAfterDelete = True
            cbBox.SelStart = Len(cbBox) 'set insertion position @ the end of string
        ElseIf KeyAscii = vbKeyReturn Then  'user select this string
            cbBox.SelStart = Len(cbBox)
            lResult = SendMessage(cbBox.hwnd, CBN_SELENDOK, 0, 0)
            AutoMatchCBBox = KeyAscii 'let caller a chance to handle "Enter"
        End If
        
    Else 'generate searching string
        If lLength = 0 Then
            strFindThis = cbBox.Text & Chr(KeyAscii) 'No selection, append it
        Else
            strFindThis = Left(cbBox.Text, lStart) & Chr(KeyAscii)
        End If
    End If
    
    If bContinueSearch Then 'need to search
        Call VBComBoBoxDroppedDown(cbBox)  'open dropdown list
        If (Not bContinueSearchAfterDelete) Then
            lResult = SendMessage(cbBox.hwnd, CB_SELECTSTRING, -1, ByVal strFindThis)
        Else
            
            lStart = lStart - 1
            If lStart = -1 Then
                'empty value
                lStart = 0
                strFindThis = ""
                
            Else
                strFindThis = Left(cbBox.Text, lStart)
                lResult = SendMessage(cbBox.hwnd, CB_SELECTSTRING, -1, ByVal strFindThis)
             End If
        End If
        If lResult = CB_ERR Then 'not found
        'if not found, act as not pressed
            'lStart = lStart - 1
            strFindThis = Left(cbBox.Text, lStart)
            lResult = SendMessage(cbBox.hwnd, CB_SELECTSTRING, -1, ByVal strFindThis)
             If lResult <> CB_ERR Then
                cbBox.SelStart = Len(strFindThis)
                cbBox.SelLength = Len(cbBox) - cbBox.SelStart
            End If
'
'            cbBox.Text = strFindThis 'set cbBox as whatever it is
'            cbBox.SelLength = 0 'no selected char(s) since not found
'            cbBox.SelStart = Len(cbBox) 'set insertion position @ the end of string
       
        Else
            'found string, highlight rest of string for user
            cbBox.SelStart = Len(strFindThis)
            cbBox.SelLength = Len(cbBox) - cbBox.SelStart
            
        End If
    End If
    On Error GoTo 0
    Exit Function
    
ErrHandle:
    'got problem, simply return whatever pass in
    Debug.Print "Failed: AutoCompleteComboBox due to : " & Err.Description
    Debug.Assert False
    AutoMatchCBBox = KeyAscii
    On Error GoTo 0
End Function

'open dorpdown list
Private Sub VBComBoBoxDroppedDown(ByRef cbBox As ComboBox)
    Call SendMessage(cbBox.hwnd, CB_SHOWDROPDOWN, Abs(True), 0)
End Sub



Private Sub cmbPhraseTemplate_KeyDown(KeyCode As Integer, Shift As Integer)
'block delete key, it causes an unwanted behaviour that allows
If KeyCode = vbKeyDelete Then
    KeyCode = 0
End If
End Sub

Private Sub cmbPhraseTemplate_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbPhraseTemplate, KeyAscii)
End Sub

Private Sub cmbPhraseTemplate_LostFocus()

    SetBackColor
    
    
End Sub


Private Sub UserControl_GotFocus()
    cmbPhraseTemplate.SetFocus
End Sub

Private Sub UserControl_Initialize()
    cmbPhraseTemplate.Left = 0
    cmbPhraseTemplate.Top = 0
    cmbPhraseTemplate.Width = UserControl.Width
    UserControl.Height = cmbPhraseTemplate.Height
End Sub

Public Sub Initialize()
    Dim i As Integer
    Dim cPhrase As clsPhraseTemplate
    cmbPhraseTemplate.List(0) = ""
    BuildCmbList
    For i = 1 To cmbPhraseTemplate.ListCount - 1
        Set cPhrase = DictPhrase.Item(cmbPhraseTemplate.List(i))
        If Trim(cPhrase.Code) = Trim(InitContent) Then
            cmbPhraseTemplate.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Public Function GetValue() As String
    GetValue = Code
End Function


Private Sub UserControl_LostFocus()
    SetBackColor
     Dim cPhrase As clsPhraseTemplate
   
    If DictPhrase.Exists(cmbPhraseTemplate.Text) Then
        Set cPhrase = DictPhrase.Item(cmbPhraseTemplate.Text)
        If cPhrase.Code <> "" Then
            Exit Sub
        End If
    End If
    cmbPhraseTemplate.Text = ""
End Sub

Private Sub UserControl_Resize()
    cmbPhraseTemplate.Left = 0
    cmbPhraseTemplate.Top = 0
    cmbPhraseTemplate.Width = UserControl.Width
    UserControl.Height = cmbPhraseTemplate.Height
End Sub

Private Sub UserControl_Show()
    cmbPhraseTemplate.Left = 0
    cmbPhraseTemplate.Top = 0
    cmbPhraseTemplate.Width = UserControl.Width
    UserControl.Height = cmbPhraseTemplate.Height
End Sub

Private Sub BuildCmbList()
    Dim SQLStr As String
    Dim cPhrase As clsPhraseTemplate
        
    SQLStr = Trim(Replace(PhraseName, "$$", ""))
'    SQLStr = "select pe.phrase_description,pe.phrase_info " & _
             "from lims_sys.phrase_entry pe, lims_sys.phrase_header ph " & _
             "where pe.phrase_id = ph.phrase_id and ph.name = '" & PhraseName & "' " & _
             "order by order_number"
             
    Set rst = Connection.Execute(SQLStr)
    While Not rst.EOF
        Set cPhrase = New clsPhraseTemplate
        cPhrase.Exception = IIf(rst("exception").Value = "T", True, False)
        cPhrase.Code = rst("code").Value
        cPhrase.Text = rst("text").Value
        cPhrase.Color = rst("color").Value
        cmbPhraseTemplate.AddItem cPhrase.Text
        DictPhrase.Add cPhrase.Text, cPhrase
        rst.MoveNext
    Wend
End Sub

Public Sub Terminate()
    cmbPhraseTemplate.Clear
    DictPhrase.RemoveAll
    Set DictPhrase = Nothing
    Set rst = Nothing
End Sub

Public Property Let Locked(ByVal vNewValue As Boolean)
    cmbPhraseTemplate.Locked = vNewValue
End Property

Public Property Let FontName(fname As String)
    cmbPhraseTemplate.Font.Name = fname
End Property

Public Property Let RightMargin(RMargin As Long)
    cmbPhraseTemplate.RightToLeft = RMargin
End Property

Public Property Get PTBHandle() As Long
    PTBHandle = cmbPhraseTemplate.hwnd
End Property
Public Property Get Rtl() As Boolean
    Rtl = cmbPhraseTemplate.RightToLeft
End Property

Public Property Let Rtl(ByVal vNewValue As Boolean)
    cmbPhraseTemplate.RightToLeft = vNewValue
End Property

Public Property Get Execption() As Boolean
    Dim cPhrase As clsPhraseTemplate
    Execption = False
    If Not DictPhrase.Exists(cmbPhraseTemplate.List(cmbPhraseTemplate.ListIndex)) Then Exit Property
    Set cPhrase = DictPhrase.Item(cmbPhraseTemplate.List(cmbPhraseTemplate.ListIndex))
    Execption = cPhrase.Exception
End Property


Public Property Get Code() As String
    Dim cPhrase As clsPhraseTemplate
    Code = ""
    If DictPhrase.Exists(cmbPhraseTemplate.Text) Then
        Set cPhrase = DictPhrase.Item(cmbPhraseTemplate.Text)
        Code = cPhrase.Code
    End If

'
'   If Not DictPhrase.Exists(cmbPhraseTemplate.List(cmbPhraseTemplate.ListIndex)) Then Exit Property
'    Set cPhrase = DictPhrase.Item(cmbPhraseTemplate.List(cmbPhraseTemplate.ListIndex))
'   Code = cPhrase.Code
End Property
Public Property Get Text() As String
    Dim cPhrase As clsPhraseTemplate
    Text = ""
    If Not DictPhrase.Exists(cmbPhraseTemplate.List(cmbPhraseTemplate.ListIndex)) Then Exit Property
    Set cPhrase = DictPhrase.Item(cmbPhraseTemplate.List(cmbPhraseTemplate.ListIndex))
    Text = cPhrase.Text
End Property

Public Property Get Color() As String
    Dim cPhrase As clsPhraseTemplate
    Color = ""

    If Not DictPhrase.Exists(cmbPhraseTemplate.List(cmbPhraseTemplate.ListIndex)) Then
        Color = vbWhite
        Exit Property
    End If

    Set cPhrase = DictPhrase.Item(cmbPhraseTemplate.List(cmbPhraseTemplate.ListIndex))
    Color = cPhrase.Color
End Property

Private Sub SetBackColor()
    cmbPhraseTemplate.BackColor = CLng(Trim(Color))
    cmbPhraseTemplate.Refresh
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oTextBox As TextBox
Attribute m_oTextBox.VB_VarHelpID = -1

Public Property Set TextBox(ByVal oTextBox As TextBox)
    Set m_oTextBox = oTextBox
End Property

Private Sub m_oTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Call OnlyNumbers(KeyAscii, m_oTextBox)
    
    If m_oTextBox.Name Like "num*" Then
        'If Not (KeyAscii > 47 And KeyAscii < 59) Then
        '    Beep
        '    MsgBox "Whole numbers only, please."
        '    KeyAscii = 0
        '    m_oTextBox.SetFocus
        '    m_oTextBox.SelLength = 0
        'End If
        
        Select Case KeyAscii
            Case 47 To 57
            Case Else
                Beep
                MsgBox "Whole numbers only, please."
                KeyAscii = 0
        End Select
        
    End If
    
    
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CastoffForm 
   Caption         =   "Castoff Macro"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   6240
   OleObjectBlob   =   "CastoffForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CastoffForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public blnCancel As Boolean

Private Sub cmdYesCastoff_Click()
    Dim blnTrimStatus As Boolean
    Dim blnDesignStatus As Boolean
    
    blnTrimStatus = False
    blnDesignStatus = False
    
    If Me.optTrim5x8 Or Me.optTrim6x9 Then
        blnTrimStatus = True
    End If
    
    If Me.optDesignLoose Or Me.optDesignAverage Or Me.optDesignTight Then
        blnDesignStatus = True
    End If
    
    If blnTrimStatus = True And blnDesignStatus = True Then
        blnCancel = False
        Me.Hide
    Else
        MsgBox "You must select one Trim Size and one Design to generate a castoff."
    End If
End Sub

Private Sub cmdNoCastoff_Click()
    blnCancel = True
    Me.Hide
End Sub

Private Sub tabPublisher_Change()

    Dim i As String
    
    i = Me.tabPublisher.SelectedItem.Caption

    Select Case i
        Case "SMP"
            optTrim5x8.Enabled = True
            optTrim6x9.Enabled = True
        
            optDesignLoose.Enabled = True
            optDesignAverage.Enabled = True
            optDesignTight.Enabled = True
            
        Case "torDOTcom"
            optTrim5x8.Enabled = True
            optTrim6x9.Enabled = False
        
            optDesignLoose.Enabled = True
            optDesignAverage.Enabled = False
            optDesignTight.Enabled = False
    End Select

End Sub

Private Sub UserForm_Initialize()

    'To ensure consistent appearance on different OS
    Dim lngHexVal As Long
    lngHexVal = &HF3F3F3

    CastoffForm.BackColor = lngHexVal
    cmdNoCastoff.BackColor = lngHexVal
    cmdYesCastoff.BackColor = lngHexVal
    fraDesign.BackColor = lngHexVal
    fraTrimSize.BackColor = lngHexVal
    labReminder.BackColor = lngHexVal
    tabPublisher.BackColor = lngHexVal
    optTrim5x8.BackColor = lngHexVal
    optTrim6x9.BackColor = lngHexVal
    optDesignLoose.BackColor = lngHexVal
    optDesignAverage.BackColor = lngHexVal
    optDesignTight.BackColor = lngHexVal
    fraMissingPages.BackColor = lngHexVal
    labMissingPages.BackColor = lngHexVal
    
    'set all option buttons to not selected
    txtMissingPages.Text = "0"
    optTrim5x8.Value = False
    optTrim6x9.Value = False
    optDesignLoose.Value = False
    optDesignAverage.Value = False
    optDesignTight.Value = False

End Sub

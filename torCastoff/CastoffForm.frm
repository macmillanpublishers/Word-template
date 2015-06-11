VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CastoffForm 
   Caption         =   "Castoff Macro"
   ClientHeight    =   4725
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
            optTrim5x8.Visible = True
            optTrim6x9.Visible = True
        
            optDesignLoose.Visible = True
            optDesignAverage.Visible = True
            optDesignTight.Visible = True
            
        Case "torDOTcom"
            optTrim5x8.Value = True
            optTrim5x8.Visible = True
            optTrim6x9.Visible = False
        
            optDesignLoose.Value = True
            optDesignLoose.Visible = True
            optDesignAverage.Visible = False
            optDesignTight.Visible = False
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

End Sub

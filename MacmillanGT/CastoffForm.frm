VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CastoffForm 
   Caption         =   "Macmillan Preliminary Castoff Form"
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   12225
   OleObjectBlob   =   "CastoffForm.frx":0000
   ShowModal       =   0   'False
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
    
    If Me.chkDesignLoose Or Me.chkDesignAverage Or Me.chkDesignTight Then
        blnDesignStatus = True
    End If
    
    If blnTrimStatus = True And blnDesignStatus = True Then
        blnCancel = False
        Me.Hide
    Else
        MsgBox "You must select one Trim Size and at least one Design to generate a castoff."
    End If
End Sub

Private Sub cmdNoCastoff_Click()
    blnCancel = True
    Me.Hide
End Sub



'Private Sub UserForm_Initialize()
'
'    'To ensure consistent appearance on different OS
'    Dim lngHexVal As Long
'    lngHexVal = &HF3F3F3
'
'    Me.BackColor = lngHexVal
'    cmdNoCastoff.BackColor = lngHexVal
'    cmdYesCastoff.BackColor = lngHexVal
'    fraDesign.BackColor = lngHexVal
'    fraTrimSize.BackColor = lngHexVal
'    labReminder.BackColor = lngHexVal
'    tabPublisher.BackColor = lngHexVal
'    optTrim5x8.BackColor = lngHexVal
'    optTrim6x9.BackColor = lngHexVal
'    chkDesignLoose.BackColor = lngHexVal
'    chkDesignAverage.BackColor = lngHexVal
'    chkDesignTight.BackColor = lngHexVal
'    fraMissingPages.BackColor = lngHexVal
'    labMissingPages.BackColor = lngHexVal
    
    'set all option buttons to not selected
'    txtMissingPages.Text = "0"
'    optTrim5x8.value = True
'    optTrim6x9.value = False
'    chkDesignLoose.value = True
'    chkDesignAverage.value = True
'    chkDesignTight.value = True
    
    'make sure text is 10 pt because sometimes it turns into 2pt?
'    fraDesign.Font.Size = 10
'    fraMissingPages.Font.Size = 10
'    fraTrimSize.Font.Size = 10

'End Sub


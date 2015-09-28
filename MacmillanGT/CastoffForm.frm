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
Public blnHelp As Boolean

Private Sub cmdYesCastoff_Click()
    ' Cancel or Help were not clicked
    blnCancel = False
    blnHelp = False
    
    Dim blnTitleStatus As Boolean
    Dim blnPubStatus As Boolean
    Dim blnTrimStatus As Boolean
    Dim blnDesignStatus As Boolean
    
    blnTitleStatus = False
    blnPubStatus = False
    blnTrimStatus = False
    blnDesignStatus = False
    
    'Has something been entered for all Title Info fields?
    If Me.txtEditor <> vbNullString And Me.txtAuthor <> vbNullString And Me.txtTitle <> vbNullString _
        And Me.txtPageCount <> vbNullString Then
            blnTitleStatus = True
    End If
    
    'Has something been selected for Publisher?
    If Me.optPubSMP Or Me.optPubTor Or Me.optPubPickup Then
        blnPubStatus = True
    End If
    
    'Has something been selected for Trim Size?
    If Me.optTrim5x8 Or Me.optTrim6x9 Then
        blnTrimStatus = True
    End If
    
    'Has something been selected for Design?
    If Me.chkDesignLoose Or Me.chkDesignAverage Or Me.chkDesignTight Then
        blnDesignStatus = True
    End If
    
    'OK if all required have been set, otherwise give a warning message.
    If blnTrimStatus = True And blnDesignStatus = True And blnPubStatus = True And blnTitleStatus = True Then
        blnCancel = False
        Me.Hide
    Else
        MsgBox "You must fill in Title Info, Publisher, Trim Size, and Design to generate a castoff."
    End If
End Sub


Private Sub cmdNoCastoff_Click()
    blnHelp = False
    blnCancel = True
    Me.Hide
End Sub


Private Sub cmdHelp_Click()
    blnCancel = False
    blnHelp = True
    Me.Hide
End Sub

Private Sub UserForm_Initialize()

    'To ensure consistent appearance on different OS
    Dim lngHexVal As Long
    Dim lngHexRed As Long
    lngHexVal = &HF3F3F3
    lngHexRed = &HC0

    Me.BackColor = lngHexVal
    labHeading.BackColor = lngHexVal
    
    fraTitleInfo.BackColor = lngHexVal
    fraTitleInfo.ForeColor = lngHexRed
    labEditor.BackColor = lngHexVal
    labAuthor.BackColor = lngHexVal
    labTitle.BackColor = lngHexVal
    labPageCount.BackColor = lngHexVal
    
    fraPublisher.BackColor = lngHexVal
    fraPublisher.ForeColor = lngHexRed
    optPubSMP.BackColor = lngHexVal
    optPubTor.BackColor = lngHexVal
    optPubPickup.BackColor = lngHexVal
    
    fraTrimSize.BackColor = lngHexVal
    fraTrimSize.ForeColor = lngHexRed
    optTrim5x8.BackColor = lngHexVal
    optTrim6x9.BackColor = lngHexVal
    
    fraDesign.BackColor = lngHexVal
    fraDesign.ForeColor = lngHexRed
    chkDesignLoose.BackColor = lngHexVal
    chkDesignAverage.BackColor = lngHexVal
    chkDesignTight.BackColor = lngHexVal
    
    fraStandard.BackColor = lngHexVal
    labChapters.BackColor = lngHexVal
    labParts.BackColor = lngHexVal
    labFrontmatter.BackColor = lngHexVal
    
    fraBackmatter.BackColor = lngHexVal
    labIndex.BackColor = lngHexVal
    labBackmatter.BackColor = lngHexVal
    
    fraNotesBib.BackColor = lngHexVal
    labUnlinkedNotes.BackColor = lngHexVal
    labNotesTK.BackColor = lngHexVal
    labBibliography.BackColor = lngHexVal
    labBiblioTK.BackColor = lngHexVal
    
    fraComplex.BackColor = lngHexVal
    labSubheads.BackColor = lngHexVal
    labTables.BackColor = lngHexVal
    labArt.BackColor = lngHexVal
    
    fraPickup.BackColor = lngHexVal
    labPrevTitle.BackColor = lngHexVal
    labPrevPageCount.BackColor = lngHexVal
    labPrevCharCount.BackColor = lngHexVal
    labAddlPgs.BackColor = lngHexVal
    
    cmdNoCastoff.BackColor = lngHexVal
    cmdYesCastoff.BackColor = lngHexVal
    cmdHelp.BackColor = lngHexVal
    
    'set all default selections
    optTrim5x8.value = True
    optTrim6x9.value = False
    chkDesignLoose.value = True
    chkDesignAverage.value = True
    chkDesignTight.value = True
    
    'make sure frame text is 10 pt because sometimes it turns into 2pt and I don't know why
    labHeading.Font.Size = 12
    fraTitleInfo.Font.Size = 10
    fraPublisher.Font.Size = 10
    fraDesign.Font.Size = 10
    fraTrimSize.Font.Size = 10
    fraStandard.Font.Size = 10
    fraBackmatter.Font.Size = 10
    fraNotesBib.Font.Size = 10
    fraComplex.Font.Size = 10
    fraPickup.Font.Size = 10

End Sub


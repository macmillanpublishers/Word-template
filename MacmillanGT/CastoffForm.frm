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
Const lngHexVal As Long = &HF3F3F3      'Background color of userform
Const lngHexRed As Long = &HC0          'Red for required sections
Const lngHexBlack As Long = &H0             'Black for non-required sections

Private m_oCollectionOfEventHandlers As Collection


Private Sub labBiblioTK_Click()

End Sub

Private Sub UserForm_Initialize()
    
    ' Create instance of TextboxEvenHandler for each control
    ' Started from http://stackoverflow.com/questions/1083603/vba-using-withevents-on-userforms
    Set m_oCollectionOfEventHandlers = New Collection

    Dim oControl As control
    For Each oControl In Me.Controls

        If TypeName(oControl) = "TextBox" Then

            Dim oEventHandler As TextBoxEventHandler
            Set oEventHandler = New TextBoxEventHandler

            Set oEventHandler.TextBox = oControl

            m_oCollectionOfEventHandlers.Add oEventHandler

        End If

    Next oControl


    'Set userform appearance to ensure consistent appearance on different OS

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
    optTrim5x8.value = False
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


Private Sub cmdYesCastoff_Click()
    ' Cancel was not clicked
    blnCancel = False
    
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
        And Me.numTxtPageCount <> vbNullString Then
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
    Else
        Me.Hide
        MsgBox "You must fill in Title Info, Publisher, Trim Size, and Design to generate a castoff."
        blnCancel = True
        Me.Show
    End If
    
    'Also all "Standard" inputs are required if SMP or Tor.com, all "Pickup" are required if "Pickup Design"
    If Me.optPubSMP Or Me.optPubTor Then
        If Me.numTxtChapters = vbNullString Or Me.numTxtParts = vbNullString Or Me.numTxtFrontmatter = vbNullString Then
            Me.Hide
            MsgBox "You must fill in Standard Items section to get a castoff."
            blnCancel = True
            Me.Show
        Else
            Me.Hide
        End If
    ElseIf Me.optPubPickup Then
        If Me.txtPrevTitle = vbNullString Or Me.numTxtPrevPageCount = vbNullString Or Me.numTxtPrevCharCount = vbNullString _
            Or Me.numTxtAddlPgs = vbNullString Then
                Me.Hide
                MsgBox "You must full in the Pickup Designs section to get a castoff."
                blnCancel = True
                Me.Show
        Else
            Me.Hide
        End If
    End If
    
    If blnCancel = False Then
        Call CastoffStart(FormInputs:=Me)
    End If
            
    
End Sub


Private Sub cmdNoCastoff_Click()
    blnCancel = True
    Unload Me
End Sub


Private Sub cmdHelp_Click()
    blnCancel = False
    
    Me.Hide
    
    Dim strHelpMessage As String
        
    strHelpMessage = "MACMILLAN PRELIMINARY CASTOFF FORM" & vbNewLine & vbNewLine & _
    "Note: These are ballpark estimates only. Characters per page are finally determined by font, font size, " & _
    "and text width." & vbNewLine & vbNewLine & _
    "This form will calculate an estimated print page count based on the manuscript file you run it on and " & _
    "the information you enter on this form." & vbNewLine & vbNewLine & _
    "You can find more detailed information about this form at <Confluence Page>, or contact " & _
    "workflows@macmillan.com if you have any questions."
    
    MsgBox strHelpMessage, vbOKOnly, "Castoff Help"
    Me.Show
    

End Sub



Private Sub optPubSMP_Click()
    fraStandard.ForeColor = lngHexRed
    fraPickup.ForeColor = lngHexBlack
End Sub


Private Sub optPubTor_Click()
    fraStandard.ForeColor = lngHexRed
    fraPickup.ForeColor = lngHexBlack
End Sub


Private Sub optPubPickup_Click()
    fraStandard.ForeColor = lngHexBlack
    fraPickup.ForeColor = lngHexRed
End Sub


 


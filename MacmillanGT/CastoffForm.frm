VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CastoffForm 
   Caption         =   "Macmillan Preliminary Castoff Form"
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   9360
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
Private cBookTitle As String
Private cAuthorName As String
Private cImprint As String



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
    labHeading2.BackColor = lngHexVal
    labHeading2.ForeColor = lngHexRed
    labHeading3.BackColor = lngHexVal
    
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
    
    fraPrintType.BackColor = lngHexVal
    fraPrintType.ForeColor = lngHexRed
    optPrintOffset.BackColor = lngHexVal
    optPrintPOD.BackColor = lngHexVal
    
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
    fraPrintType.Font.Size = 10
    
    ' ===== FOR TESTING ONLY =================
    ' ===== COMMENT OUT FOR PRODUCTION =======
    ' txtEditor.value = "Editor Name"
    ' txtAuthor.value = "Author Name"
    ' txtTitle.value = "Book Title"
    ' numTxtPageCount = "224"
    ' optPubTor.value = True
    ' optTrim5x8.value = True
    ' numTxtChapters.value = "10"
    ' numTxtParts.value = "2"
    ' numTxtFrontmatter.value = "14"
    
    ' Get metadata from doc if it's styled
    Me.BookTitle = GetText("Titlepage Book Title (tit)")
    Me.AuthorName = GetText("Titlepage Author Name (au)")
    Me.Imprint = GetText("Titlepage Imprint Line (imp)")

End Sub


Private Sub cmdYesCastoff_Click()
    ' Cancel was not clicked
    blnCancel = False
    
    Dim blnTitleStatus As Boolean
    Dim blnPubStatus As Boolean
    Dim blnPrintStatus As Boolean
    Dim blnTrimStatus As Boolean
    Dim blnDesignStatus As Boolean
    
    blnTitleStatus = False
    blnPubStatus = False
    blnPrintStatus = False
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
    
    ' Has something been selected for Print Type?
    If Me.optPrintOffset Or Me.optPrintPOD Then
        blnPrintStatus = True
    End If
    
    'Has something been selected for Trim Size?
    If Me.optTrim5x8 Or Me.optTrim6x9 Then
        blnTrimStatus = True
    End If
    
    'Has something been selected for Design?
    If Me.optPubPickup Then
        blnDesignStatus = True
    Else
        If Me.chkDesignLoose Or Me.chkDesignAverage Or Me.chkDesignTight Then
            blnDesignStatus = True
        End If
    End If
    
    'OK if all required have been set, otherwise give a warning message.
    If blnTrimStatus = True And blnDesignStatus = True And blnPubStatus = True And blnTitleStatus = True _
        And blnPrintStatus = True Then
        blnCancel = False
    Else
        Me.Hide
        MsgBox "You must fill in Title Info, Publisher, Print Type, Trim Size, and Design to generate a castoff."
        blnCancel = True
        Me.Show
        Exit Sub
    End If
    
    ' Check that scheduled page count is multiple of 16
    If Me.optPrintOffset And Me.numTxtPageCount <> vbNullString Then
        If Me.numTxtPageCount Mod 16 > 0 Then
            Me.Hide
            MsgBox "Scheduled page count must be a multiple of 16."
            blnCancel = True
            Me.Show
            Exit Sub
        End If
    End If
    
    'Also all "Standard" inputs are required if SMP or Tor.com, all "Pickup" are required if "Pickup Design"
    If Me.optPubSMP Or Me.optPubTor Then
        If Me.numTxtChapters = vbNullString Or Me.numTxtParts = vbNullString Or Me.numTxtFrontmatter = vbNullString Then
            Me.Hide
            MsgBox "You must fill in the Standard Items section to get a castoff."
            blnCancel = True
            Me.Show
            Exit Sub
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
                Exit Sub
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
        
    Dim strHelpMessage As String
        
    strHelpMessage = "MACMILLAN PRELIMINARY CASTOFF FORM" & vbNewLine & vbNewLine & _
    "This form will calculate an estimated print page count based on the manuscript file you run it on and " & _
    "the information you enter on this form." & vbNewLine & vbNewLine & _
    "Note: These are ballpark estimates only. Characters per page are finally determined by font, font size, " & _
    "and text width." & vbNewLine & vbNewLine & _
    "You can find more detailed information about this form at <Confluence Page>, or contact " & _
    "workflows@macmillan.com if you have any questions."
    
    MsgBox strHelpMessage, vbOKOnly, "Castoff Help"

End Sub


Private Sub optPubSMP_Click()
    fraStandard.ForeColor = lngHexRed
    fraPickup.ForeColor = lngHexBlack
    fraDesign.ForeColor = lngHexRed
    
    optPrintPOD.value = False
    optPrintOffset.value = True
    
    optTrim5x8.value = False
    optTrim6x9.Enabled = True
    
    chkDesignLoose.value = True
    chkDesignLoose.Enabled = True
    chkDesignAverage.value = True
    chkDesignAverage.Enabled = True
    chkDesignTight.value = True
    chkDesignTight.Enabled = True
    
End Sub


Private Sub optPubTor_Click()
    fraStandard.ForeColor = lngHexRed
    fraPickup.ForeColor = lngHexBlack
    fraDesign.ForeColor = lngHexRed
    
    optPrintPOD.value = True
    optPrintOffset.value = False
    
    optTrim5x8.value = True
    optTrim6x9.Enabled = False
    
    chkDesignLoose.value = False
    chkDesignLoose.Enabled = False
    chkDesignAverage.value = True
    chkDesignAverage.Enabled = True
    chkDesignTight.value = False
    chkDesignTight.Enabled = False

End Sub


Private Sub optPubPickup_Click()
    fraStandard.ForeColor = lngHexBlack
    fraPickup.ForeColor = lngHexRed
    fraDesign.ForeColor = lngHexBlack
        
    optPrintPOD.value = False
    optPrintOffset.value = True
    
    optTrim5x8.value = False
    optTrim6x9.Enabled = True
    
    chkDesignLoose.value = False
    chkDesignLoose.Enabled = False
    chkDesignAverage.value = False
    chkDesignAverage.Enabled = False
    chkDesignTight.value = False
    chkDesignTight.Enabled = False
    
End Sub


' ============= Now we're creating some properties for the CastoffForm inputs to get from the text if styled ========

Public Property Let BookTitle(value As String)
' This procedure is executed when the BookTitle Property is set.
    cBookTitle = value
    Me.txtTitle = cBookTitle
End Property


Public Property Get BookTitle() As String
' This is executed when the user tries to access the property.
    BookTitle = cBookTitle
End Property


Public Property Let AuthorName(value As String)
    cAuthorName = value
    Me.txtAuthor = cAuthorName
End Property


Public Property Get AuthorName() As String
    AuthorName = cAuthorName
End Property


Public Property Let Imprint(value As String)
    cImprint = value
    If InStr(1, cImprint, "Martin") > 0 Then  ' InStr because cImprint = Me.optPubSmp.Caption fails if apostrophe is curly
        Me.optPubSMP.value = True
    ElseIf cImprint = Me.optPubTor.Caption Then
        Me.optPubTor.value = True
    End If
End Property


Public Property Get Imprint() As String
    Imprint = cImprint
End Property

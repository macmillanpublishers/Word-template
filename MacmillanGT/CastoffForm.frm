VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CastoffForm 
   Caption         =   "Macmillan Preliminary Castoff Form"
   ClientHeight    =   10785
   ClientLeft      =   45
   ClientTop       =   -1380
   ClientWidth     =   9720
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

' For TextBoxEventHandler, which throws a warning if user enters non-numerals in number boxes
Private m_oCollectionOfEventHandlers As Collection

' For custom properties
Private cPublisherCodeString As String
Private cTrimSizeString As String
Private cTrimIndexLong As Long
Private cImprintString As String
Private cPublisherString As String

' ============= First we're creating some properties for the CastoffForm inputs ========
' ============= Inputs that are just a text entry already have a property (value) ====
' ============= but option buttons don't have a property for which one was selected ==

' This is for the way the book will be printed
' Will be validated that user has selected one in cmdYes_Click event
Public Property Get PrintType() As String
' Get value for property from option buttons
    If optPrintOffset.value = True Then
        PrintType = optPrintOffset.Caption
    ElseIf optPrintPOD.value = True Then
        PrintType = optPrintPOD.Caption
    End If
End Property

Public Property Let PrintType(strPrintType As String)
' When value is assigned to property, select the correct option button
' NOTE: when assigning values, use option button caption, not literal
    If strPrintType = Me.optPrintOffset.Caption Then
        Me.optPrintOffset.value = True
    ElseIf strPrintType = Me.optPrintPOD.Caption Then
        Me.optPrintPOD.value = True
    End If
End Property

' This one is for the string of the trim size
' Will also be validated that user has selected one in cmdYes_Click event
Public Property Get TrimSize() As String
' Get Trim Size from option buttons
    If Me.optTrim5x8.value = True Then
        TrimSize = Me.optTrim5x8.Caption
    ElseIf Me.optTrim6x9.value = True Then
        TrimSize = Me.optTrim6x9.Caption
    End If
End Property

Public Property Let TrimSize(strTrimSize As String)
' If value assigned to property, select the correct option button
' NOTE: when assigning values, use option button caption, not literal
    Select Case strTrimSize
        Case Me.optTrim5x8.Caption
            Me.optTrim5x8.value = True
        Case Me.optTrim6x9.Caption
            Me.optTrim6x9.value = True
        End Select
End Property
' This is for the index of the trim size in the castoff CSV/array


Public Property Let Imprint(strImprintValue As String)
' This is executed when the Imprint property is set.
' This holds the imprint as listed on the title page, if doc is styled.
' If the document is styled correctly, and the styled Imprint Line matches one of the
' Publisher options buttons, that button will be selected. If they don't match, the
' correct imprint can still be put in the output file

    cImprintString = strImprintValue
    
    If InStr(1, cImprintString, "Martin") > 0 Then  ' InStr because cImprint = Me.optPubSmp.Caption fails if apostrophe is curly
        Me.optPubSMP = True
    ElseIf cImprintString = Me.optPubTor.Caption Then
        Me.optPubTor = True
    End If
    
End Property

Public Property Get Imprint() As String
' This is executed when the user tries to access the property.
    Imprint = cImprintString
End Property


Public Property Let PublisherCode(value As String)
    cPublisherCodeString = value
End Property

Private Sub UserForm_Initialize()
    
    ' Create instance of TextboxEvenHandler for each control
    ' Which throws a warning if anything other than numerals are entered
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
    fraTitleInfo.Font.Bold = True
    fraPublisher.Font.Size = 10
    fraPublisher.Font.Bold = True
    fraDesign.Font.Size = 10
    fraDesign.Font.Bold = True
    fraTrimSize.Font.Size = 10
    fraTrimSize.Font.Bold = True
    fraStandard.Font.Size = 10
    fraStandard.Font.Bold = True
    fraBackmatter.Font.Size = 10
    fraBackmatter.Font.Bold = True
    fraNotesBib.Font.Size = 10
    fraNotesBib.Font.Bold = True
    fraComplex.Font.Size = 10
    fraComplex.Font.Bold = True
    fraPickup.Font.Size = 10
    fraPickup.Font.Bold = True
    fraPrintType.Font.Size = 10
    fraPrintType.Font.Bold = True
    
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
    txtTitle = GetText("Titlepage Book Title (tit)")
    txtAuthor = GetText("Titlepage Author Name (au)")
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
    
    ' Set some properties based on userform imputs
    
    
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




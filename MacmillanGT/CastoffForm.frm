VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CastoffForm 
   Caption         =   "Macmillan Preliminary Castoff Form"
   ClientHeight    =   11010
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
Private cImprintString As String

' ============= First we're creating some properties for the CastoffForm inputs ========
' ============= Inputs that are just a text entry already have a property (value) ====
' ============= but option buttons don't have a property for which one was selected ==

' This is for the way the book will be printed
' Will be validated that user has selected one in cmdYes_Click event
Public Property Get PrintType() As String
' Get value for property from option buttons
' Which may have been set by the user or the Property Let function
    If optPrintOffset.value = True Then
        PrintType = optPrintOffset.Caption
    ElseIf optPrintPOD.value = True Then
        PrintType = optPrintPOD.Caption
    End If
End Property

Public Property Let PrintType(strPrintType As String)
' When value is assigned to property, select the correct option button
' This will be picked up from Property Get via the option buttons
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
' Which may have been set by the user or the Property Let function
    If Me.optTrim5x8.value = True Then
        TrimSize = Me.optTrim5x8.Caption
    ElseIf Me.optTrim6x9.value = True Then
        TrimSize = Me.optTrim6x9.Caption
    End If
End Property

Public Property Let TrimSize(strTrimSize As String)
' If value assigned to property, select the correct option button
' This will be picked up from Property Get via the option buttons
' NOTE: when assigning values, use option button caption, not literal
    Select Case strTrimSize
        Case Me.optTrim5x8.Caption
            Me.optTrim5x8.value = True
        Case Me.optTrim6x9.Caption
            Me.optTrim6x9.value = True
        End Select
End Property

' This is for the index number of the trim size in the castoff CSV/array
Public Property Get TrimIndex() As Long
' Get trim based on option buttons
' No Property Let because this is read only, no values allowed outside these two
    If Me.optTrim5x8.value = True Then
        TrimIndex = 0
    ElseIf Me.optTrim6x9.value = True Then
        TrimIndex = 1
    End If
End Property

' This holds the actual imprint as listed on the title page, if doc is styled.
' If not styled, uses option buttons
' not calling "publisher" because it won't always match publisher option buttons
Public Property Let Imprint(strImprintValue As String)
' Store the passed value, and if it matches one of the Publisher option buttons, select it
' NOTE: use GetText() function in SharedMacros module to get from styled doc

    cImprintString = strImprintValue
    
    If InStr(1, cImprintString, "Martin") > 0 Then  ' InStr because cImprint = Me.optPubSmp.Caption fails if apostrophe is curly
        Me.optPubSMP = True
    ElseIf cImprintString = Me.optPubTor.Caption Then
        Me.optPubTor = True
    End If
    
End Property

Public Property Get Imprint() As String
' This is executed when the user tries to access the property.
' If the imprint has already been set, use that
' If it hasn't, get it from the option buttons
' The default is SMP

    If cImprintString <> "" Then
        Imprint = cImprintString
    ElseIf Me.optPubTor.value = True Then
        Imprint = Me.optPubTor.Caption
    Else
        Imprint = Me.optPubSMP.Caption
    End If
    
End Property

' This holds the short code for the publishers that is used in the castoff file name
Public Property Get PublisherCode() As String
' Get code based on options buttons
' If publisher matches styled imprint line, the buttons have already been set
' Read only so no property let

' Default is SMP.
' Add new publishers here as ElseIf
    If Me.optPubTor.value = True Then
        PublisherCode = "torDOTcom"
    Else
        PublisherCode = "SMP"
    End If
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
    Me.labHeading.BackColor = lngHexVal
    Me.labHeading2.BackColor = lngHexVal
    Me.labHeading2.ForeColor = lngHexRed
    Me.labHeading3.BackColor = lngHexVal
    
    Me.fraTitleInfo.BackColor = lngHexVal
    Me.fraTitleInfo.ForeColor = lngHexRed
    Me.labEditor.BackColor = lngHexVal
    Me.labAuthor.BackColor = lngHexVal
    Me.labTitle.BackColor = lngHexVal
    Me.labPageCount.BackColor = lngHexVal
    
    Me.fraPublisher.BackColor = lngHexVal
    Me.fraPublisher.ForeColor = lngHexRed
    Me.optPubSMP.BackColor = lngHexVal
    Me.optPubTor.BackColor = lngHexVal
    
    Me.fraPrintType.BackColor = lngHexVal
    Me.fraPrintType.ForeColor = lngHexRed
    Me.optPrintOffset.BackColor = lngHexVal
    Me.optPrintPOD.BackColor = lngHexVal
    
    Me.fraTrimSize.BackColor = lngHexVal
    Me.fraTrimSize.ForeColor = lngHexRed
    Me.optTrim5x8.BackColor = lngHexVal
    Me.optTrim6x9.BackColor = lngHexVal
    
    Me.fraDesign.BackColor = lngHexVal
    Me.fraDesign.ForeColor = lngHexRed
    Me.chkDesignLoose.BackColor = lngHexVal
    Me.chkDesignAverage.BackColor = lngHexVal
    Me.chkDesignTight.BackColor = lngHexVal
    Me.chkDesignPickup.BackColor = lngHexVal
    
    Me.fraStandard.BackColor = lngHexVal
    Me.labChapters.BackColor = lngHexVal
    Me.labParts.BackColor = lngHexVal
    Me.labFrontmatter.BackColor = lngHexVal
    
    Me.fraBackmatter.BackColor = lngHexVal
    Me.labIndex.BackColor = lngHexVal
    Me.labBackmatter.BackColor = lngHexVal
    
    Me.fraNotesBib.BackColor = lngHexVal
    Me.labUnlinkedNotes.BackColor = lngHexVal
    Me.labNotesTK.BackColor = lngHexVal
    Me.labBibliography.BackColor = lngHexVal
    Me.labBiblioTK.BackColor = lngHexVal
    
    Me.fraComplex.BackColor = lngHexVal
    Me.labSubheads.BackColor = lngHexVal
    Me.labTables.BackColor = lngHexVal
    Me.labArt.BackColor = lngHexVal
    
    Me.fraPickup.BackColor = lngHexVal
    Me.labPrevTitle.BackColor = lngHexVal
    Me.labPrevPageCount.BackColor = lngHexVal
    Me.labPrevCharCount.BackColor = lngHexVal
    Me.labAddlPgs.BackColor = lngHexVal
    
    Me.cmdNoCastoff.BackColor = lngHexVal
    Me.cmdYesCastoff.BackColor = lngHexVal
    Me.cmdHelp.BackColor = lngHexVal
    
    'make sure frame text is 10 pt because sometimes it turns into 2pt when saved on Mac
    labHeading.Font.Size = 12
    Me.fraTitleInfo.Font.Size = 10
    Me.fraTitleInfo.Font.Bold = True
    Me.fraPublisher.Font.Size = 10
    Me.fraPublisher.Font.Bold = True
    Me.fraDesign.Font.Size = 10
    Me.fraDesign.Font.Bold = True
    Me.fraTrimSize.Font.Size = 10
    Me.fraTrimSize.Font.Bold = True
    Me.fraStandard.Font.Size = 10
    Me.fraStandard.Font.Bold = True
    Me.fraBackmatter.Font.Size = 10
    Me.fraBackmatter.Font.Bold = True
    Me.fraNotesBib.Font.Size = 10
    Me.fraNotesBib.Font.Bold = True
    Me.fraComplex.Font.Size = 10
    Me.fraComplex.Font.Bold = True
    Me.fraPickup.Font.Size = 10
    Me.fraPickup.Font.Bold = True
    Me.fraPrintType.Font.Size = 10
    Me.fraPrintType.Font.Bold = True
    
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
    Me.txtTitle = GetText("Titlepage Book Title (tit)")
    Me.txtAuthor = GetText("Titlepage Author Name (au)")
    ' This one selects the option button if it matches the imprint line
    Me.Imprint = GetText("Titlepage Imprint Line (imp)")
    
    'set all default selections
    Me.optTrim5x8.value = False
    Me.optTrim6x9.value = False
    Me.chkDesignLoose.value = True
    Me.chkDesignAverage.value = True
    Me.chkDesignTight.value = True
    Me.chkDesignPickup.value = False

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
    If Me.chkDesignLoose Or Me.chkDesignAverage Or Me.chkDesignTight Or Me.chkDesignPickup Then
        blnDesignStatus = True
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
    If Me.chkDesignPickup.value = False Then
        If Me.numTxtChapters = vbNullString Or Me.numTxtParts = vbNullString Or Me.numTxtFrontmatter = vbNullString Then
            Me.Hide
            MsgBox "You must fill in the Standard Items section to get a castoff."
            blnCancel = True
            Me.Show
            Exit Sub
        Else
            Me.Hide
        End If
    ElseIf Me.chkDesignPickup.value = True Then
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
    Else ' this won't happen but let's have it anyway
        Me.Hide
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
    ' Make required sections' headings red, nonrequired black
    Me.fraStandard.ForeColor = lngHexRed
    Me.fraPickup.ForeColor = lngHexBlack
    Me.fraDesign.ForeColor = lngHexRed
    
    ' Default Print Type to Offset for SMP (user can choose POD though)
    Me.PrintType = Me.optPrintOffset.Caption
    
    ' Default for trim is NEITHER selected
    Me.optTrim5x8.value = False
    Me.optTrim6x9.value = False
    
    ' MAke sure both are enabled though
    Me.optTrim5x8.Enabled = True
    Me.optTrim6x9.Enabled = True
    
    ' Make sure each design is enabled AND checked (but not pickup)
    Me.chkDesignLoose.value = True
    Me.chkDesignLoose.Enabled = True
    
    Me.chkDesignAverage.value = True
    Me.chkDesignAverage.Enabled = True
    
    Me.chkDesignTight.value = True
    Me.chkDesignTight.Enabled = True
    
    Me.chkDesignPickup.value = False
    Me.chkDesignPickup.Enabled = True
    
End Sub


Private Sub optPubTor_Click()
    ' Make required sections' headings red, nonrequired black
    Me.fraStandard .ForeColor = lngHexRed
    Me.fraPickup.ForeColor = lngHexBlack
    Me.fraDesign.ForeColor = lngHexRed
    
    ' Only print type for Tor.com is POD (for now)
    Me.PrintType = Me.optPrintPOD.Caption
    
    ' Only trim size is 5 x 8, disable other
    Me.optTrim5x8.Enabled = True
    Me.optTrim5x8.value = True
    
    Me.optTrim6x9.value = False
    Me.optTrim6x9.Enabled = False
    
    ' only design allowed is average
    Me.chkDesignLoose.value = False
    Me.chkDesignLoose.Enabled = False
    
    Me.chkDesignAverage.value = True
    Me.chkDesignAverage.Enabled = True
    
    Me.chkDesignTight.value = False
    Me.chkDesignTight.Enabled = False
        
    Me.chkDesignPickup.value = False
    Me.chkDesignPickup.Enabled = False

End Sub


Private Sub chkDesignPickup_Click()
    ' required sections' heading colors
    Me.fraStandard.ForeColor = lngHexBlack
    Me.fraPickup.ForeColor = lngHexRed
    Me.fraDesign.ForeColor = lngHexBlack
        
    ' I guess you could do POD but we'll default to Offset
    Me.PrintType = Me.optPrintOffset.Caption
    
    ' enable both trims but don't pick either
    Me.optTrim5x8.Enabled = True
    Me.optTrim5x8.value = False
    
    Me.optTrim6x9.Enabled = True
    Me.optTrim6x9.value = False
    
    ' don't pick any designs, we're getting design from prev title
    Me.chkDesignLoose.value = False
    Me.chkDesignLoose.Enabled = False
    
    Me.chkDesignAverage.value = False
    Me.chkDesignAverage.Enabled = False
    
    Me.chkDesignTight.value = False
    Me.chkDesignTight.Enabled = False
    
End Sub




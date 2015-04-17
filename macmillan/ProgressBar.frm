VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Macro Progress"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9345
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========================Progress bar================================.
'bits and pieces taken from https://strugglingtoexcel.wordpress.com/2014/09/22/class-progress-bar-excel-vba/

'Private variables for storing the Main Properties of the Object

'Assigned to Caption Property of the Form (ProgressBar).
Private cFormTitle As String

'Assigned to Caption Property of LabelCaption control; displays message to user.
Private cStatusMessage As String

'Assigned to Width property of LabelProgress control.
Private cBarWidth As Single

'Assigned to the Caption property of FrameProgress control.
Private cPercentComplete As Single


'##########################################################################
'Class Events
'##########################################################################

'This Procedure is Run when the Form in Initiated
Private Sub UserForm_Initialize()
'Set Default Values for all the Variables

cFormTitle = "Progress Bar"
cStatusMessage = "Ready"
cPercentComplete = 0
cBarWidth = cPercentComplete * Me.FrameProgress.Width

Me.Caption = cFormTitle
Me.LabelCaption.Caption = cStatusMessage
Me.FrameProgress.Caption = Format(cPercentComplete, "0%")
Me.LabelProgress.Width = cBarWidth
End Sub


'##########################################################################
'Properties
'##########################################################################

'This Procedure is Executed when the Title Propoerty is Set
Public Property Let Title(value As String)

    'Proceed if the user did not send a blank string
'If Not value = vbNullString Then
    'Initialize the private class variable
    cFormTitle = value
    'Update the Form's title if it has already been loaded
    'If Not Me Is Nothing Then
        'Do Events makes sure the rest of your macro keeps running
        'DoEvents
        Me.Caption = cFormTitle
        'Me.Repaint  'Supposedly less computationally expensive as DoEvents?
   ' End If
'End If


End Property

'This Procedure lets the User try to access the Title Property.
Public Property Get Title() As String
Title = cFormTitle
End Property



'-----------------------------------------------------------------------
'This Procedure is Executed when the StatusMessage Property is Set.
Public Property Let StatusMessage(value As String)

cStatusMessage = value
Me.LabelCaption = cStatusMessage

End Property

'This Procedure lets the User try to access the StatusMessage Property.
Public Property Get StatusMessage() As String
StatusMessage = cStatusMessage
End Property



'-------------------------------------------------------------------------
'This Procedure is Executed when the Percent Property is Set.
Public Property Let Percent(value As Single)

cPercentComplete = Format(value, "0%")
Me.FrameProgress.Caption = cPercentComplete

End Property

'This Procedure lets the User try to access the Percent Property.
Public Property Get Percent() As Single
Percent = cPercentComplete
End Property


'-------------------------------------------------------------------------
'This Procedure is Executed when the BarWidth Property is Set.
Public Property Let BarWidth(value As Single)

cBarWidth = value * Me.FrameProgress.Width
Me.LabelProgress.Width = cBarWidth

End Property

'This Procedure lets the User try to access the Percent Property.
Public Property Get BarWidth() As Single
BarWidth = cBarWidth
End Property

'##########################################################################
'Public Methods
'##########################################################################

Public Sub Progress(sglPctComplete As Single, strStatus As String)

Me.StatusMessage (strStatus)
Me.Percent (sglPctComplete)
Me.BarWidth (sglPctComplete)
Me.Repaint

End Sub


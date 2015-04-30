VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Macro Progress"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9345
   OleObjectBlob   =   "ProgressBar.frx":0000
   ShowModal       =   0   'False
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

Me.Title = cFormTitle
Me.LabelCaption.Caption = cStatusMessage
Me.FrameProgress.Caption = cPercentComplete
Me.LabelProgress.Width = cBarWidth

End Sub


'##########################################################################
'Properties
'##########################################################################

'This Procedure is Executed when the Title Property is Set
Public Property Let Title(value As String)
cFormTitle = value
Me.Caption = cFormTitle
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
cPercentComplete = Format(value * 100, "0")
Me.FrameProgress.Caption = cPercentComplete & "%"
End Property

'This Procedure lets the User try to access the Percent Property.
Public Property Get Percent() As Single
Percent = cPercentComplete
End Property

'-------------------------------------------------------------------------
'This Procedure is Executed when the BarWidth Property is Set.
Public Property Let BarWidth(value As Single)
cBarWidth = value * (Me.FrameProgress.Width - 21)
Me.LabelProgress.Width = cBarWidth
End Property

'This Procedure lets the User try to access the Percent Property.
Public Property Get BarWidth() As Single
BarWidth = cBarWidth
End Property

'##########################################################################
'Public Methods
'##########################################################################
Public Sub Increment(sglPctComplete As Single, strStatus As String)

Me.StatusMessage = strStatus
Me.Percent = sglPctComplete
Me.BarWidth = sglPctComplete
Me.Repaint

End Sub


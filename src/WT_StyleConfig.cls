VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WT_StyleConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Public Properties
Private prop_VersatileStyles As Collection
Private prop_SectionStartStyles As Collection

' Private Properties
Private prop_VbaStyleConfig As Dictionary


' *****************************************************************************
'       PUBLIC PROPERTIES
' *****************************************************************************

' ===== VersatileStyles =======================================================
' Collection of styles that can appear within a contiguous block. This downloads
' the JSON and loads it to a collection object (using JsonConverter).

Public Property Get VersatileStyles() As Collection
  If prop_VersatileStyles Is Nothing Then
    Set prop_VersatileStyles = VbaStyleConfig("versatileparas")
  End If
  Set VersatileStyles = prop_VersatileStyles
End Property

' ===== SectionStartStyles ====================================================
' Collection of "section-start" styles. This downloads the JSON and loads it to
' a Collection object.

Public Property Get SectionStartStyles() As Collection
  If prop_SectionStartStyles Is Nothing Then
    Set prop_SectionStartStyles = VbaStyleConfig("sectionstarts")
  End If
  Set SectionStartStyles = prop_SectionStartStyles
End Property


' *****************************************************************************
'       PRIVATE PROPERTIES
' *****************************************************************************

' ===== VbaStyleConfig ========================================================
' Dictionary to store data from "vba_style_config.json" file. Read-only, but
' downloads file (or uses local version if available and download fails).

Private Property Get VbaStyleConfig() As Dictionary
  If prop_VbaStyleConfig Is Nothing Then
    Dim strJsonPath As String
    Set prop_VbaStyleConfig = New Dictionary
    
    strJsonPath = SharedFileInstaller.DownloadJson("vba_style_config.json")
    Set prop_VbaStyleConfig = ClassHelpers.ReadJson(strJsonPath)
  End If
  Set VbaStyleConfig = prop_VbaStyleConfig
End Property



' *****************************************************************************
'       PUBLIC METHODS
' *****************************************************************************

' ===== IsVersatileStyle ======================================================
' Tests if a style is in the VersatileStyles collection

' PARAMS:
' StyleName: string style name to test

' RETURNS:
' Boolean

Public Function IsVersatileStyle(StyleName As String) As Boolean
  IsVersatileStyle = Utils.ExistsInCollection(Me.VersatileStyles, StyleName)
End Function

' ===== IsSectionStartStyle ===================================================
' Tests if a style is in the SectionStartStyles collection

' PARAMS:
' StyleName: string style name to test

' RETURNS:
' Boolean

Public Function IsSectionStartStyle(StyleName As String) As Boolean
  IsSectionStartStyle = Utils.ExistsInCollection(Me.SectionStartStyles, StyleName)
End Function

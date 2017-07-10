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
Private prop_AllowedBlankStyles As Collection
Private prop_VersatileStylesByIndex As Collection
Private prop_MacmillanStyles As Dictionary
Private prop_BookmakerStyles As Collection


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

' ===== AllowedBlankStyles ====================================================
' Collection of styles that are allowed to be blank (no text). VbaStyleConfig
' downloads the source JSON file.

Public Property Get AllowedBlankStyles() As Collection
  If prop_AllowedBlankStyles Is Nothing Then
    Set prop_AllowedBlankStyles = VbaStyleConfig("allowedblankparas")
  End If
  Set AllowedBlankStyles = prop_AllowedBlankStyles
End Property

' ===== VersatileStylesByIndex ================================================
' Collection of paragraph index of each style in current document that has a
' versatile style applied

Public Property Get VersatileStylesByIndex() As Collection
  If prop_VersatileStylesByIndex Is Nothing Then
    Set prop_VersatileStylesByIndex = New Collection
    Dim varStyle As Variant
    For Each varStyle In Me.VersatileStyles
      ClassHelpers.FindStyleIndex CStr(varStyle), prop_VersatileStylesByIndex
    Next varStyle
  End If
  Set VersatileStylesByIndex = prop_VersatileStylesByIndex
End Property


' ===== MacmillanStyles =======================================================
' Hard coded for now, move to style config json.

Public Property Get MacmillanStyles() As Dictionary
  If prop_MacmillanStyles Is Nothing Then
    Set prop_MacmillanStyles = New Dictionary
    prop_MacmillanStyles.Add "pagebreak", "Page Break (pb)"
    
  End If
  Set MacmillanStyles = prop_MacmillanStyles
End Property

' ===== BookmakerStyles =======================================================
' Collection of styles allowed in Bookmaker

Public Property Get BookmakerStyles() As Collection
  If prop_BookmakerStyles Is Nothing Then
    Set prop_BookmakerStyles = VbaStyleConfig("bookmakerstyles")
  End If
  Set BookmakerStyles = prop_BookmakerStyles
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

' ===== VersatileStyleExistsByIndex ===========================================
' Returns True if the style applied to the paragraph of activeDoc w/ that index
' number is a VersatileStyle

Public Function VersatileStyleExistsByIndex(ParagraphIndex As Long) As Boolean
  VersatileStyleExistsByIndex = Utils.ExistsInCollection(Me.VersatileStylesByIndex, _
    ParagraphIndex)
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

' ===== IsBookmakerStyle ======================================================
' Tests if a style is in the BookmakerStyles collection

Public Function IsBookmakerStyle(StyleName As String) As Boolean
  IsBookmakerStyle = Utils.ExistsInCollection(Me.BookmakerStyles, StyleName)
End Function

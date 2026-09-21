Attribute VB_Name = "modAlbumInit"
Option Explicit

' ============================================================================
' ALBUM DA COPA - MODULO DE INICIALIZACAO AUTOMATICA
' Carrega os UserForms ao abrir o workbook
' ============================================================================

Public Sub InitializeAlbum()
    ' Called from Workbook_Open event
    On Error Resume Next
    
    ' Check if UserForms exist
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    ' Force load of UserForms into memory
    Dim frmAlbum As Object
    Dim frmCompare As Object
    
    Set frmAlbum = CreateObject("frmAlbum")
    Set frmCompare = CreateObject("frmCompare")
    
    ' Unload immediately - just to compile/load them
    Unload frmAlbum
    Unload frmCompare
    
    On Error GoTo 0
End Sub

Public Sub ShowAlbum()
    ' Public sub to show the main album form
    On Error Resume Next
    frmAlbum.Show
    On Error GoTo 0
End Sub

Public Sub ShowCompare()
    ' Public sub to show the compare form
    On Error Resume Next
    frmCompare.Show
    On Error GoTo 0
End Sub
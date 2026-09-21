Attribute VB_Name = "frmCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ============================================================================
' ALBUM DA COPA - USERFORM DE COMPARACAO (2 JOGADORES LADO A LADO)
' Rise Kujikawa - Investigation Team AI
' ============================================================================

Private m_playerData As Object
Private m_isLoading As Boolean

Private Sub UserForm_Initialize()
    SetupForm
    LoadDropdowns
    LoadPlayerData
End Sub

Private Sub SetupForm()
    Me.Caption = "ALBUM DA COPA - COMPARAR JOGADORES"
    Me.Width = 1200
    Me.Height = 800
    Me.StartUpPosition = 1
    Me.BackColor = RGB(13, 27, 42)
    
    StyleControls
End Sub

Private Sub StyleControls()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Label" Then
            ctrl.ForeColor = RGB(255, 215, 0)
            ctrl.BackStyle = fmBackStyleTransparent
            ctrl.Font.Name = "Calibri"
            ctrl.Font.Size = 10
        ElseIf TypeName(ctrl) = "ComboBox" Or TypeName(ctrl) = "ListBox" Then
            ctrl.BackColor = RGB(255, 255, 255)
            ctrl.ForeColor = RGB(27, 42, 74)
            ctrl.Font.Name = "Calibri"
            ctrl.Font.Size = 10
            ctrl.BorderStyle = fmBorderStyleSingle
        ElseIf TypeName(ctrl) = "CommandButton" Then
            ctrl.BackColor = RGB(27, 42, 74)
            ctrl.ForeColor = RGB(255, 215, 0)
            ctrl.Font.Name = "Calibri"
            ctrl.Font.Size = 10
            ctrl.Font.Bold = True
            ctrl.BorderStyle = fmBorderStyleSingle
        ElseIf TypeName(ctrl) = "Frame" Then
            ctrl.ForeColor = RGB(255, 215, 0)
            ctrl.BackColor = RGB(13, 27, 42)
            ctrl.Font.Name = "Calibri"
            ctrl.Font.Size = 11
            ctrl.Font.Bold = True
        ElseIf TypeName(ctrl) = "TextBox" Then
            ctrl.BackColor = RGB(255, 255, 255)
            ctrl.ForeColor = RGB(27, 42, 74)
            ctrl.Font.Name = "Calibri"
            ctrl.Font.Size = 10
            ctrl.BorderStyle = fmBorderStyleSingle
        End If
    Next ctrl
End Sub

Private Sub LoadDropdowns()
    Dim k As Variant
    cmbPlayer1.Clear
    cmbPlayer2.Clear
    
    For Each k In m_playerData.Keys
        cmbPlayer1.AddItem k
        cmbPlayer2.AddItem k
    Next k
    
    cmbCup1.Clear
    cmbCup2.Clear
    cmbCup1.AddItem "2010"
    cmbCup1.AddItem "2014"
    cmbCup1.AddItem "2018"
    cmbCup1.AddItem "2022"
    cmbCup1.AddItem "2026"
    cmbCup2.AddItem "2010"
    cmbCup2.AddItem "2014"
    cmbCup2.AddItem "2018"
    cmbCup2.AddItem "2022"
    cmbCup2.AddItem "2026"
End Sub

Private Sub LoadPlayerData()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim playerDict As Object
    Set playerDict = CreateObject("Scripting.Dictionary")
    
    Set ws = ThisWorkbook.Sheets("dim_players_final")
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets("Players")
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        Dim name As String
        name = ws.Cells(i, 2).Value
        If name <> "" Then
            Dim data As Object
            Set data = CreateObject("Scripting.Dictionary")
            data("Name") = name
            data("Country") = ws.Cells(i, 3).Value
            data("Position") = ws.Cells(i, 4).Value
            data("Age") = ws.Cells(i, 5).Value
            data("Height") = ws.Cells(i, 6).Value
            data("Weight") = ws.Cells(i, 7).Value
            data("Foot") = ws.Cells(i, 8).Value
            data("MarketValue") = ws.Cells(i, 9).Value
            data("Rating") = ws.Cells(i, 10).Value
            data("Matches") = ws.Cells(i, 11).Value
            data("Minutes") = ws.Cells(i, 12).Value
            data("Goals") = ws.Cells(i, 13).Value
            data("Assists") = ws.Cells(i, 14).Value
            data("xG") = ws.Cells(i, 15).Value
            data("Shots") = ws.Cells(i, 16).Value
            data("ShotsOnTarget") = ws.Cells(i, 17).Value
            data("Dribbles") = ws.Cells(i, 18).Value
            data("PassAccuracy") = ws.Cells(i, 19).Value
            data("Tackles") = ws.Cells(i, 20).Value
            data("Interceptions") = ws.Cells(i, 21).Value
            data("Clearances") = ws.Cells(i, 22).Value
            data("Fouls") = ws.Cells(i, 23).Value
            data("YellowCards") = ws.Cells(i, 24).Value
            data("RedCards") = ws.Cells(i, 25).Value
            data("Saves") = ws.Cells(i, 26).Value
            data("GoalsConceded") = ws.Cells(i, 27).Value
            data("CleanSheets") = ws.Cells(i, 28).Value
            playerDict(name) = data
        End If
    Next i
    
    Set m_playerData = playerDict
End Sub

Private Sub cmbPlayer1_Change()
    If m_isLoading Then Exit Sub
    UpdatePlayer1Display
End Sub

Private Sub cmbPlayer2_Change()
    If m_isLoading Then Exit Sub
    UpdatePlayer2Display
End Sub

Private Sub cmbCup1_Change()
    If m_isLoading Then Exit Sub
    UpdatePlayer1Display
End Sub

Private Sub cmbCup2_Change()
    If m_isLoading Then Exit Sub
    UpdatePlayer2Display
End Sub

Private Sub UpdatePlayer1Display()
    Dim playerName As String
    playerName = cmbPlayer1.Value
    If playerName = "" Then Exit Sub
    
    If Not m_playerData.Exists(playerName) Then Exit Sub
    
    Dim data As Object
    Set data = m_playerData(playerName)
    
    ' Frame 1 caption
    fraPlayer1.Caption = "JOGADOR 1: " & UCase(playerName)
    
    ' Info
    lblP1Name.Caption = data("Name")
    lblP1Country.Caption = data("Country")
    lblP1Cup.Caption = cmbCup1.Value
    lblP1Pos.Caption = data("Position")
    lblP1Age.Caption = data("Age") & " anos"
    lblP1HW.Caption = data("Height") & " cm / " & data("Weight") & " kg"
    lblP1Foot.Caption = data("Foot")
    lblP1Value.Caption = Format(data("MarketValue"), "#,##0.0") & " M€"
    lblP1Rating.Caption = Format(data("Rating"), "0.0")
    
    ' Stats
    UpdatePlayerStats fraPlayer1, data, "1"
    
    ' Per 90
    UpdatePlayerPer90 fraPlayer1, data, "1"
End Sub

Private Sub UpdatePlayer2Display()
    Dim playerName As String
    playerName = cmbPlayer2.Value
    If playerName = "" Then Exit Sub
    
    If Not m_playerData.Exists(playerName) Then Exit Sub
    
    Dim data As Object
    Set data = m_playerData(playerName)
    
    ' Frame 2 caption
    fraPlayer2.Caption = "JOGADOR 2: " & UCase(playerName)
    
    ' Info
    lblP2Name.Caption = data("Name")
    lblP2Country.Caption = data("Country")
    lblP2Cup.Caption = cmbCup2.Value
    lblP2Pos.Caption = data("Position")
    lblP2Age.Caption = data("Age") & " anos"
    lblP2HW.Caption = data("Height") & " cm / " & data("Weight") & " kg"
    lblP2Foot.Caption = data("Foot")
    lblP2Value.Caption = Format(data("MarketValue"), "#,##0.0") & " M€"
    lblP2Rating.Caption = Format(data("Rating"), "0.0")
    
    ' Stats
    UpdatePlayerStats fraPlayer2, data, "2"
    
    ' Per 90
    UpdatePlayerPer90 fraPlayer2, data, "2"
End Sub

Private Sub UpdatePlayerStats(frame As Frame, data As Object, suffix As String)
    Dim mins As Double
    mins = data("Minutes")
    If mins = 0 Then mins = 1
    
    Dim lblMatches As Control
    Dim lblMinutes As Control
    Dim lblGoals As Control
    Dim lblAssists As Control
    Dim lblxG As Control
    Dim lblShots As Control
    Dim lblShotsOT As Control
    Dim lblDribbles As Control
    Dim lblPassAcc As Control
    Dim lblTackles As Control
    Dim lblIntercepts As Control
    Dim lblClearances As Control
    Dim lblFouls As Control
    Dim lblYellow As Control
    Dim lblRed As Control
    Dim lblCleanSheets As Control
    Dim lblSaves As Control
    Dim lblGoalsConc As Control
    
    Set lblMatches = frame.Controls("lblP" & suffix & "Matches")
    Set lblMinutes = frame.Controls("lblP" & suffix & "Minutes")
    Set lblGoals = frame.Controls("lblP" & suffix & "Goals")
    Set lblAssists = frame.Controls("lblP" & suffix & "Assists")
    Set lblxG = frame.Controls("lblP" & suffix & "xG")
    Set lblShots = frame.Controls("lblP" & suffix & "Shots")
    Set lblShotsOT = frame.Controls("lblP" & suffix & "ShotsOT")
    Set lblDribbles = frame.Controls("lblP" & suffix & "Dribbles")
    Set lblPassAcc = frame.Controls("lblP" & suffix & "PassAcc")
    Set lblTackles = frame.Controls("lblP" & suffix & "Tackles")
    Set lblIntercepts = frame.Controls("lblP" & suffix & "Intercepts")
    Set lblClearances = frame.Controls("lblP" & suffix & "Clearances")
    Set lblFouls = frame.Controls("lblP" & suffix & "Fouls")
    Set lblYellow = frame.Controls("lblP" & suffix & "Yellow")
    Set lblRed = frame.Controls("lblP" & suffix & "Red")
    Set lblCleanSheets = frame.Controls("lblP" & suffix & "CleanSheets")
    Set lblSaves = frame.Controls("lblP" & suffix & "Saves")
    Set lblGoalsConc = frame.Controls("lblP" & suffix & "GoalsConc")
    
    If Not lblMatches Is Nothing Then
        lblMatches.Caption = data("Matches")
        lblMinutes.Caption = data("Minutes")
        lblGoals.Caption = data("Goals")
        lblAssists.Caption = data("Assists")
        lblxG.Caption = Format(data("xG"), "0.00")
        lblShots.Caption = data("Shots")
        lblShotsOT.Caption = data("ShotsOnTarget")
        lblDribbles.Caption = data("Dribbles")
        lblPassAcc.Caption = Format(data("PassAccuracy"), "0.0") & "%"
        lblTackles.Caption = data("Tackles")
        lblIntercepts.Caption = data("Interceptions")
        lblClearances.Caption = data("Clearances")
        lblFouls.Caption = data("Fouls")
        lblYellow.Caption = data("YellowCards")
        lblRed.Caption = data("RedCards")
        
        If data("Position") = "GK" Then
            lblCleanSheets.Caption = data("CleanSheets")
            lblSaves.Caption = data("Saves")
            lblGoalsConc.Caption = data("GoalsConceded")
        Else
            lblCleanSheets.Caption = "-"
            lblSaves.Caption = "-"
            lblGoalsConc.Caption = "-"
        End If
    End If
End Sub

Private Sub UpdatePlayerPer90(frame As Frame, data As Object, suffix As String)
    Dim mins As Double
    mins = data("Minutes")
    If mins = 0 Then mins = 1
    
    Dim lblP90Goals As Control
    Dim lblP90Assists As Control
    Dim lblP90xG As Control
    Dim lblP90Shots As Control
    Dim lblP90Passes As Control
    Dim lblP90PassAcc As Control
    Dim lblP90Dribbles As Control
    Dim lblP90Tackles As Control
    Dim lblP90Intercepts As Control
    Dim lblP90Fouls As Control
    
    Set lblP90Goals = frame.Controls("lblP" & suffix & "P90Goals")
    Set lblP90Assists = frame.Controls("lblP" & suffix & "P90Assists")
    Set lblP90xG = frame.Controls("lblP" & suffix & "P90xG")
    Set lblP90Shots = frame.Controls("lblP" & suffix & "P90Shots")
    Set lblP90Passes = frame.Controls("lblP" & suffix & "P90Passes")
    Set lblP90PassAcc = frame.Controls("lblP" & suffix & "P90PassAcc")
    Set lblP90Dribbles = frame.Controls("lblP" & suffix & "P90Dribbles")
    Set lblP90Tackles = frame.Controls("lblP" & suffix & "P90Tackles")
    Set lblP90Intercepts = frame.Controls("lblP" & suffix & "P90Intercepts")
    Set lblP90Fouls = frame.Controls("lblP" & suffix & "P90Fouls")
    
    If Not lblP90Goals Is Nothing Then
        lblP90Goals.Caption = Format(data("Goals") / mins * 90, "0.00")
        lblP90Assists.Caption = Format(data("Assists") / mins * 90, "0.00")
        lblP90xG.Caption = Format(data("xG") / mins * 90, "0.00")
        lblP90Shots.Caption = Format(data("Shots") / mins * 90, "0.0")
        lblP90Passes.Caption = "0.0"
        lblP90PassAcc.Caption = Format(data("PassAccuracy"), "0.0")
        lblP90Dribbles.Caption = Format(data("Dribbles") / mins * 90, "0.0")
        lblP90Tackles.Caption = Format(data("Tackles") / mins * 90, "0.0")
        lblP90Intercepts.Caption = Format(data("Interceptions") / mins * 90, "0.0")
        lblP90Fouls.Caption = Format(data("Fouls") / mins * 90, "0.0")
    End If
End Sub

Private Sub btnCompareRadar_Click()
    ' Show radar chart comparison
    MsgBox "Radar Chart Comparison would open here with both players overlayed", vbInformation, "Album da Copa"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
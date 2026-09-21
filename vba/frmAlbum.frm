Attribute VB_Name = "frmAlbum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ============================================================================
' ALBUM DA COPA - USERFORM PRINCIPAL (FIGURINHA DO JOGADOR)
' Rise Kujikawa - Investigation Team AI
' ============================================================================

Private m_currentPlayer As String
Private m_currentCup As String
Private m_currentCountry As String
Private m_currentPosition As String
Private m_playerData As Collection
Private m_isLoading As Boolean

' Radar Chart axes by position
Private m_radarAxes As Object
Private m_radarMax As Object

Private Sub UserForm_Initialize()
    SetupForm()
    LoadDropdowns
    SetupRadarChart
    LoadPlayerData
End Sub

Private Sub SetupForm()
    Me.Caption = "ALBUM DA COPA DEFINITIVO 2010-2026 - Figurinha do Jogador"
    Me.Width = 1000
    Me.Height = 750
    Me.StartUpPosition = 1 ' Center screen
    Me.BackColor = RGB(13, 27, 42) ' Dark navy
    
    ' Style all controls
    StyleControls
End Sub

Private Sub StyleControls()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Label" Then
            ctrl.ForeColor = RGB(255, 215, 0) ' Gold
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
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Load from Figurinha sheet
    Set ws = ThisWorkbook.Sheets("Figurinha")
    
    ' Jogador dropdown (from data validation)
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    For i = 5 To lastRow
        If ws.Cells(i, 3).Value <> "" And ws.Cells(i, 3).Value <> "Buscar Jogador:" Then
            dict(ws.Cells(i, 3).Value) = 1
        End If
    Next i
    
    cmbPlayer.Clear
    For Each k In dict.Keys
        cmbPlayer.AddItem k
    Next k
    Set dict = Nothing
    
    ' Copa dropdown
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    For i = 5 To lastRow
        If ws.Cells(i, 7).Value <> "" And ws.Cells(i, 7).Value <> "Copa:" Then
            dict(ws.Cells(i, 7).Value) = 1
        End If
    Next i
    
    cmbCup.Clear
    For Each k In dict.Keys
        cmbCup.AddItem k
    Next k
    Set dict = Nothing
    
    ' Pais dropdown
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    For i = 5 To lastRow
        If ws.Cells(i, 9).Value <> "" And ws.Cells(i, 9).Value <> "Pais:" Then
            dict(ws.Cells(i, 9).Value) = 1
        End If
    Next i
    
    cmbCountry.Clear
    For Each k In dict.Keys
        cmbCountry.AddItem k
    Next k
    Set dict = Nothing
    
    ' Posicao dropdown
    cmbPosition.Clear
    cmbPosition.AddItem "GK"
    cmbPosition.AddItem "CB"
    cmbPosition.AddItem "RB"
    cmbPosition.AddItem "LB"
    cmbPosition.AddItem "CDM"
    cmbPosition.AddItem "CM"
    cmbPosition.AddItem "CAM"
    cmbPosition.AddItem "RW"
    cmbPosition.AddItem "LW"
    cmbPosition.AddItem "ST"
End Sub

Private Sub SetupRadarChart()
    ' Initialize radar axes for each position
    Set m_radarAxes = CreateObject("Scripting.Dictionary")
    Set m_radarMax = CreateObject("Scripting.Dictionary")
    
    ' GK: Defesas, Clean Sheets, Jogo c/ Pes, Saidas, Penaltis, Consistencia
    m_radarAxes.Add "GK", Array("Defesas/90", "Clean Sheets %", "Passes/90", "Saidas/90", "Penaltis Def.", "Consistencia")
    m_radarMax.Add "GK", Array(10, 100, 50, 3, 5, 100)
    
    ' DEF: Desarmes, Interceptacoes, Cortes, Jogo Aereo, Passes, Disciplina
    m_radarAxes.Add "CB", Array("Desarmes/90", "Intercept/90", "Cortes/90", "Venc. Aereos %", "Passes %", "Faltas/90 (Inv)")
    m_radarMax.Add "CB", Array(5, 3, 8, 100, 100, 3)
    m_radarAxes.Add "RB", Array("Desarmes/90", "Intercept/90", "Cortes/90", "Venc. Aereos %", "Passes %", "Faltas/90 (Inv)")
    m_radarMax.Add "RB", Array(5, 3, 8, 100, 100, 3)
    m_radarAxes.Add "LB", Array("Desarmes/90", "Intercept/90", "Cortes/90", "Venc. Aereos %", "Passes %", "Faltas/90 (Inv)")
    m_radarMax.Add "LB", Array(5, 3, 8, 100, 100, 3)
    
    ' MID: Passes, Progressao, Criacao, Defesa, Resistencia, Versatilidade
    m_radarAxes.Add "CDM", Array("Passes/90", "Passes Longos %", "Desarmes/90", "Intercept/90", "Min/Jogo", "Posicoes")
    m_radarMax.Add "CDM", Array(80, 100, 5, 3, 90, 5)
    m_radarAxes.Add "CM", Array("Passes/90", "Passes Prog.", "Chances Criadas", "Desarmes/90", "Min/Jogo", "G+A/90")
    m_radarMax.Add "CM", Array(80, 10, 3, 3, 90, 1)
    m_radarAxes.Add "CAM", Array("Passes/90", "Passes Chave", "Chances Criadas", "Finaliz/90", "Dribles/90", "G+A/90")
    m_radarMax.Add "CAM", Array(70, 4, 4, 3, 4, 1)
    
    ' FWD: Finalizacao, xG, Movimento, Jogo Aereo, Criacao, Clutch
    m_radarAxes.Add "RW", Array("Finaliz/90", "xG/90", "Gols/90", "Venc. Aereos %", "Assist/90", "Gols Decisivos")
    m_radarMax.Add "RW", Array(4, 0.6, 0.8, 100, 0.3, 5)
    m_radarAxes.Add "LW", Array("Finaliz/90", "xG/90", "Gols/90", "Venc. Aereos %", "Assist/90", "Gols Decisivos")
    m_radarMax.Add "LW", Array(4, 0.6, 0.8, 100, 0.3, 5)
    m_radarAxes.Add "ST", Array("Finaliz/90", "xG/90", "Gols/90", "Venc. Aereos %", "Assist/90", "Gols Decisivos")
    m_radarMax.Add "ST", Array(4, 0.6, 1.0, 100, 0.3, 8)
End Sub

Private Sub LoadPlayerData()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim playerDict As Object
    Set playerDict = CreateObject("Scripting.Dictionary")
    
    Set ws = ThisWorkbook.Sheets("dim_players_final")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("Players")
    End If
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        Dim name As String
        name = ws.Cells(i, 2).Value ' Name column
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

Private Sub cmbPlayer_Change()
    If m_isLoading Then Exit Sub
    UpdatePlayerDisplay
End Sub

Private Sub cmbCup_Change()
    If m_isLoading Then Exit Sub
    UpdatePlayerDisplay
End Sub

Private Sub cmbCountry_Change()
    If m_isLoading Then Exit Sub
    FilterPlayersByCountry
End Sub

Private Sub cmbPosition_Change()
    If m_isLoading Then Exit Sub
    FilterPlayersByPosition
End Sub

Private Sub FilterPlayersByCountry()
    Dim country As String
    country = cmbCountry.Value
    If country = "" Then Exit Sub
    
    m_isLoading = True
    cmbPlayer.Clear
    
    Dim k As Variant
    For Each k In m_playerData.Keys
        Dim data As Object
        Set data = m_playerData(k)
        If data("Country") = country Then
            cmbPlayer.AddItem k
        End If
    Next k
    m_isLoading = False
End Sub

Private Sub FilterPlayersByPosition()
    Dim pos As String
    pos = cmbPosition.Value
    If pos = "" Then Exit Sub
    
    m_isLoading = True
    cmbPlayer.Clear
    
    Dim k As Variant
    For Each k In m_playerData.Keys
        Dim data As Object
        Set data = m_playerData(k)
        If data("Position") = pos Then
            cmbPlayer.AddItem k
        End If
    Next k
    m_isLoading = False
End Sub

Private Sub UpdatePlayerDisplay()
    Dim playerName As String
    playerName = cmbPlayer.Value
    If playerName = "" Then Exit Sub
    
    If Not m_playerData.Exists(playerName) Then Exit Sub
    
    Dim data As Object
    Set data = m_playerData(playerName)
    
    m_currentPlayer = playerName
    m_currentCup = cmbCup.Value
    m_currentCountry = data("Country")
    m_currentPosition = data("Position")
    
    ' Update info labels
    lblName.Caption = data("Name")
    lblCountry.Caption = data("Country")
    lblCup.Caption = m_currentCup
    lblPosition.Caption = data("Position")
    lblAge.Caption = data("Age") & " anos"
    lblHeightWeight.Caption = data("Height") & " cm / " & data("Weight") & " kg"
    lblFoot.Caption = data("Foot")
    lblValue.Caption = Format(data("MarketValue"), "#,##0.0") & " M€"
    lblRating.Caption = Format(data("Rating"), "0.0")
    
    ' Update stats for selected cup (simplified - would query fact tables)
    UpdateCupStats data
    
    ' Update radar chart
    UpdateRadarChart data
    
    ' Update per 90
    UpdatePer90 data
    
    ' Update history
    UpdateHistory data
    
    ' Update percentiles
    UpdatePercentiles data
End Sub

Private Sub UpdateCupStats(data As Object)
    ' Simplified - would query fact_matches and fact_events for the specific cup
    Dim mins As Double
    mins = data("Minutes")
    
    If mins > 0 Then
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

Private Sub UpdateRadarChart(data As Object)
    Dim pos As String
    pos = data("Position")
    
    If Not m_radarAxes.Exists(pos) Then pos = "CM"
    
    Dim axes As Variant
    Dim maxVals As Variant
    axes = m_radarAxes(pos)
    maxVals = m_radarMax(pos)
    
    Dim mins As Double
    mins = data("Minutes")
    If mins = 0 Then mins = 1
    
    Dim values(1 To 6) As Double
    
    Select Case pos
        Case "GK"
            values(1) = Min(data("Saves") / mins * 90, maxVals(0))
            values(2) = Min(data("CleanSheets") / data("Matches") * 100, maxVals(1))
            values(3) = Min(data("PassAccuracy"), maxVals(2))
            values(4) = Min(0, maxVals(3)) ' placeholder
            values(5) = Min(0, maxVals(4)) ' placeholder
            values(6) = data("Rating") * 10
        Case "CB", "RB", "LB"
            values(1) = Min(data("Tackles") / mins * 90, maxVals(0))
            values(2) = Min(data("Interceptions") / mins * 90, maxVals(1))
            values(3) = Min(data("Clearances") / mins * 90, maxVals(2))
            values(4) = 50 ' Aerial % placeholder
            values(5) = data("PassAccuracy")
            values(6) = Max(0, maxVals(5) - data("Fouls") / mins * 90)
        Case "CDM"
            values(1) = Min(data("PassAccuracy") / 100 * 80, maxVals(0))
            values(2) = 50 ' Long passes %
            values(3) = Min(data("Tackles") / mins * 90, maxVals(2))
            values(4) = Min(data("Interceptions") / mins * 90, maxVals(3))
            values(5) = Min(data("Minutes") / data("Matches"), maxVals(4))
            values(6) = 3 ' Positions
        Case "CM"
            values(1) = Min(data("PassAccuracy") / 100 * 80, maxVals(0))
            values(2) = 5 ' Progressive passes
            values(3) = Min((data("Goals") + data("Assists")) / mins * 90, maxVals(2))
            values(4) = Min(data("Tackles") / mins * 90, maxVals(3))
            values(5) = Min(data("Minutes") / data("Matches"), maxVals(4))
            values(6) = Min((data("Goals") + data("Assists")) / mins * 90, maxVals(5))
        Case "CAM"
            values(1) = Min(data("PassAccuracy") / 100 * 70, maxVals(0))
            values(2) = 2 ' Key passes
            values(3) = 2 ' Chances created
            values(4) = Min(data("Shots") / mins * 90, maxVals(3))
            values(5) = Min(data("Dribbles") / mins * 90, maxVals(4))
            values(6) = Min((data("Goals") + data("Assists")) / mins * 90, maxVals(5))
        Case "RW", "LW", "ST"
            values(1) = Min(data("Shots") / mins * 90, maxVals(0))
            values(2) = Min(data("xG") / mins * 90, maxVals(1))
            values(3) = Min(data("Goals") / mins * 90, maxVals(2))
            values(4) = 30 ' Aerial % placeholder
            values(5) = Min(data("Assists") / mins * 90, maxVals(4))
            values(6) = 3 ' Clutch goals
    End Select
    
    ' Draw radar on chart control (would need a chart control or custom drawing)
    DrawRadarChart axes, values, maxVals
End Sub

Private Sub DrawRadarChart(axes As Variant, values As Variant, maxVals As Variant)
    ' This would draw on a chart control or use custom drawing
    ' For now, update labels
    Dim i As Integer
    For i = 0 To 5
        Dim lblAxis As Control
        Dim lblVal As Control
        Set lblAxis = Me.Controls("lblAxis" & i + 1)
        Set lblVal = Me.Controls("lblVal" & i + 1)
        If Not lblAxis Is Nothing Then
            lblAxis.Caption = axes(i)
            lblVal.Caption = Format(values(i + 1), "0.0")
        End If
    Next i
End Sub

Private Sub UpdatePer90(data As Object)
    Dim mins As Double
    mins = data("Minutes")
    If mins = 0 Then mins = 1
    
    lblP90Goals.Caption = Format(data("Goals") / mins * 90, "0.00")
    lblP90Assists.Caption = Format(data("Assists") / mins * 90, "0.00")
    lblP90xG.Caption = Format(data("xG") / mins * 90, "0.00")
    lblP90Shots.Caption = Format(data("Shots") / mins * 90, "0.0")
    lblP90Passes.Caption = "0.0" ' placeholder
    lblP90PassAcc.Caption = Format(data("PassAccuracy"), "0.0")
    lblP90Dribbles.Caption = Format(data("Dribbles") / mins * 90, "0.0")
    lblP90Tackles.Caption = Format(data("Tackles") / mins * 90, "0.0")
    lblP90Intercepts.Caption = Format(data("Interceptions") / mins * 90, "0.0")
    lblP90Fouls.Caption = Format(data("Fouls") / mins * 90, "0.0")
End Sub

Private Sub UpdateHistory(data As Object)
    ' Would query all cups for this player
    ' Placeholder
    Dim i As Integer
    For i = 1 To 5
        Dim lblHist As Control
        Set lblHist = Me.Controls("lblHist" & i)
        If Not lblHist Is Nothing Then
            lblHist.Caption = "Copa " & (2010 + (i-1)*4) & " - " & data("Country") & " - " & data("Age") & " anos"
        End If
    Next i
End Sub

Private Sub UpdatePercentiles(data As Object)
    ' Calculate percentiles vs position peers
    ' Placeholder
End Sub

Private Sub btnPrev_Click()
    ' Navigate to previous player in list
    Dim idx As Long
    idx = cmbPlayer.ListIndex
    If idx > 0 Then
        m_isLoading = True
        cmbPlayer.ListIndex = idx - 1
        m_isLoading = False
        UpdatePlayerDisplay
    End If
End Sub

Private Sub btnNext_Click()
    ' Navigate to next player in list
    Dim idx As Long
    idx = cmbPlayer.ListIndex
    If idx < cmbPlayer.ListCount - 1 Then
        m_isLoading = True
        cmbPlayer.ListIndex = idx + 1
        m_isLoading = False
        UpdatePlayerDisplay
    End If
End Sub

Private Sub btnCompare_Click()
    frmCompare.Show
End Sub

Private Sub btnOpenPBIX_Click()
    ' Open Power BI file if exists
    Dim pbixPath As String
    pbixPath = ThisWorkbook.Path & "\Album_Copa_Definitivo.pbix"
    If Dir(pbixPath) <> "" Then
        Call Shell("cmd /c start """" """ & pbixPath & """", vbNormalFocus)
    Else
        MsgBox "Arquivo Power BI nao encontrado: " & pbixPath, vbExclamation, "Album da Copa"
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
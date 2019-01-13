Attribute VB_Name = "Log"
Private Const cstCurrentGameSheet = "CURRENT GAME"
Private Const cstCurrentTurnsTable = "CURRENT_TURNS_DATA"
Private Const cstGamesSheet = "GAMES TABLE"
Private Const cstGamesTable = "GAMES_DATA"
Private Const cstTurnsSheet = "TURNS TABLE"
Private Const cstTurnsTable = "TURNS_DATA"


Public Sub InitDataInterface()

    If Not SheetExists(cstCurrentGameSheet) Then
        Call Create_CURRENT_GAME_SHEET
    End If
    
    If Not SheetExists(cstTurnsSheet) Then
        Call Create_TURNS_TABLE_SHEET
    End If

    If Not SheetExists(cstGamesSheet) Then
        Call Create_GAMES_TABLE_SHEET
    End If
    
    If SheetExists("BOARD") Then
        Worksheets("BOARD").Select
    End If
    
End Sub

Public Sub Create_GAMES_TABLE_SHEET()
Dim sh As Worksheet
Dim table As ListObject

    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = cstGamesSheet
    End With

    Set sh = Sheets(cstGamesSheet)
    
    sh.Cells.VerticalAlignment = xlTop
    sh.Cells.HorizontalAlignment = xlLeft
    
    sh.Range("A1").Value = "ID"
    sh.Range("B1").Value = "White player"
    sh.Range("C1").Value = "Black player"
    sh.Range("D1").Value = "Game date"
    sh.Range("E1").Value = "Winner"
    
    Set table = Sheets(cstGamesSheet).ListObjects.Add(xlSrcRange, Range("A1:E1"), , xlYes)
    table.Name = cstGamesTable

End Sub

Public Sub Create_CURRENT_GAME_SHEET()
Dim sh As Worksheet
Dim table As ListObject

    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = cstCurrentGameSheet
    End With
    
    Set sh = Sheets(cstCurrentGameSheet)
    
    sh.Cells.VerticalAlignment = xlTop
    sh.Cells.HorizontalAlignment = xlLeft
    
    sh.Range("A1").Value = "Turn"
    sh.Range("B1").Value = "Turn color"
    sh.Range("C1").Value = "Queen move"
    sh.Range("D1").Value = "Queen appears"
    sh.Range("E1").Value = "Pawn jumped"
    sh.Range("F1").Value = "Turn duration"
    sh.Range("G1").Value = "Board initial state"
    sh.Range("H1").Value = "Board final state"
    
    Set table = Sheets(cstCurrentGameSheet).ListObjects.Add(xlSrcRange, Range("A1:H1"), , xlYes)
    table.Name = cstCurrentTurnsTable
    
End Sub

Public Sub Create_TURNS_TABLE_SHEET()
Dim sh As Worksheet
Dim table As ListObject

    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = cstTurnsSheet
    End With
    
    Set sh = Sheets(cstTurnsSheet)
    
    sh.Cells.VerticalAlignment = xlTop
    sh.Cells.HorizontalAlignment = xlLeft
    
    sh.Range("A1").Value = "ID"
    sh.Range("B1").Value = "Game ID"
    sh.Range("C1").Value = "Turn"
    sh.Range("D1").Value = "Turn color"
    sh.Range("E1").Value = "Queen move"
    sh.Range("F1").Value = "Queen appears"
    sh.Range("G1").Value = "Pawn jumped"
    sh.Range("H1").Value = "Turn duration"
    sh.Range("I1").Value = "Board initial state"
    sh.Range("J1").Value = "Board final state"
    
    Set table = Sheets(cstTurnsSheet).ListObjects.Add(xlSrcRange, Range("A1:J1"), , xlYes)
    table.Name = cstTurnsTable
    
End Sub

Public Function CreateNewEntryAndGetID(pSheetName As String, pTableName As String) As Integer
Dim Tbl As ListObject
Dim NewRow As ListRow
Dim ID As Integer

    Set Tbl = Worksheets(pSheetName).ListObjects(pTableName)
    Set NewRow = Tbl.ListRows.Add(AlwaysInsert:=True)
    
    With Sheets(pSheetName)
        ID = .Range("A" & .Rows.Count).End(xlUp).Row - 1
    End With
    
    NewRow.Range.Cells(1, 1).Value = CStr(ID)
    CreateNewEntryAndGetID = ID
    
End Function


Public Sub CG_UpdateTurnColor(pID As Integer, pColor As EColor)
Dim rangeToUpdate As Range
    
    Set rangeToUpdate = GetRangeFromTable(cstCurrentGameSheet, cstCurrentTurnsTable, "Turn color", pID)
    rangeToUpdate.Value = EnumString(pColor)
    
End Sub

Public Sub CG_UpdateQueenMove(pID As Integer, pQueenMove As Boolean)
Dim rangeToUpdate As Range

    Set rangeToUpdate = GetRangeFromTable(cstCurrentGameSheet, cstCurrentTurnsTable, "Queen move", pID)
    rangeToUpdate.Value = Abs(Int(pQueenMove))
    
End Sub

Public Sub CG_UpdateQueenAppears(pID As Integer, pQueenAppears As Boolean)
Dim rangeToUpdate As Range

    Set rangeToUpdate = GetRangeFromTable(cstCurrentGameSheet, cstCurrentTurnsTable, "Queen appears", pID)
    rangeToUpdate.Value = Abs(Int(pQueenAppears))
    
End Sub

Public Sub CG_UpdatePawnJumped(pID As Integer, pPawnJumped As Boolean)
Dim rangeToUpdate As Range
    
    Set rangeToUpdate = GetRangeFromTable(cstCurrentGameSheet, cstCurrentTurnsTable, "Pawn jumped", pID)
    rangeToUpdate.Value = Abs(Int(pPawnJumped))

End Sub

Public Sub CG_UpdateTurnDuration(pTurnDuration As Single, Optional pID As Integer)
Dim rangeToUpdate As Range

    If pID = 0 Then
        Set rangeID = Cells(ThisWorkbook.Sheets(cstCurrentGameSheet).ListObjects(cstCurrentTurnsTable).ListRows().Count, 1)
        pID = rangeID.Row
    End If

    Set rangeToUpdate = GetRangeFromTable(cstCurrentGameSheet, cstCurrentTurnsTable, "Turn duration", pID)
    rangeToUpdate.Value = CStr(Round(pTurnDuration))
    
End Sub

Public Sub CG_UpdateBoardInitialState(pID As Integer, pSnapshot As String)
Dim rangeToUpdate As Range
    
    Set rangeToUpdate = GetRangeFromTable(cstCurrentGameSheet, cstCurrentTurnsTable, "Board initial state", pID)
    rangeToUpdate.Value = CStr(pSnapshot)

End Sub

Public Sub CG_UpdateBoardFinalState(pID As Integer, pSnapshot As String)
Dim rangeToUpdate As Range

    Set rangeToUpdate = GetRangeFromTable(cstCurrentGameSheet, cstCurrentTurnsTable, "Board final state", pID)
    rangeToUpdate.Value = CStr(pSnapshot)

End Sub

Public Sub CG_InsertNewTurn(pColor As EColor, pQueenMove As Boolean, pQueenAppears As Boolean, pPawnJumped As Boolean, pInitBoard As String, pFinalBoard As String)
Dim ID As Integer
Dim shName As String
Dim tblName As String

    ID = Log.CreateNewEntryAndGetID(cstCurrentGameSheet, cstCurrentTurnsTable)
    
    Call Log.CG_UpdateTurnColor(ID, pColor)
    Call Log.CG_UpdateQueenMove(ID, pQueenMove)
    Call Log.CG_UpdateQueenAppears(ID, pQueenAppears)
    Call Log.CG_UpdatePawnJumped(ID, pPawnJumped)
    Call Log.CG_UpdateBoardFinalState(ID, pFinalBoard)
    Call Log.CG_UpdateBoardInitialState(ID, pInitBoard)
    
    ThisWorkbook.Sheets(cstCurrentGameSheet).Rows(ID + 1).RowHeight = 12.75
    
End Sub


Public Sub TT_UpdateGameID(ByVal pID As Integer, ByVal pGameId As Integer)
Dim rangeToUpdate As Range

    Set rangeToUpdate = GetRangeFromTable(cstTurnsSheet, cstTurnsTable, "Game ID", pID)
    rangeToUpdate.Value = CStr(pGameId)

End Sub

Public Sub TT_InsertTurns(ByVal pGameId As Integer)
Dim ID As Integer
Dim turnRangeData As ListRows
Dim turn As ListRow
Dim turnsRangeData

    Set turnRangeData = Sheets(cstCurrentGameSheet).ListObjects(cstCurrentTurnsTable).ListRows

    For Each turn In turnRangeData
        ID = Log.CreateNewEntryAndGetID(cstTurnsSheet, cstTurnsTable)
        Call Log.TT_UpdateGameID(ID, pGameId)
        turn.Range.Copy Destination:=ThisWorkbook.Worksheets(cstTurnsSheet).Cells(ID + 1, 3)
        
        ThisWorkbook.Sheets(cstTurnsSheet).Rows(ID + 1).RowHeight = 12.75
    Next turn
    
End Sub

Public Sub GT_UpdateWhitePlayer(pID As Integer, pBotName As String)
Dim rangeToUpdate As Range

    Set rangeToUpdate = GetRangeFromTable(cstGamesSheet, cstGamesTable, "White player", pID)
    rangeToUpdate.Value = pBotName

End Sub

Public Sub GT_UpdateBlackPlayer(pID As Integer, pBotName As String)
Dim rangeToUpdate As Range

    Set rangeToUpdate = GetRangeFromTable(cstGamesSheet, cstGamesTable, "Black player", pID)
    rangeToUpdate.Value = pBotName

End Sub

Public Sub GT_UpdateWinner(pID As Integer, pBotName As String)
Dim rangeToUpdate As Range

    Set rangeToUpdate = GetRangeFromTable(cstGamesSheet, cstGamesTable, "Winner", pID)
    rangeToUpdate.Value = pBotName

End Sub

Public Sub GT_UpdateGameDate(pID As Integer, pDate As Date)
Dim rangeToUpdate As Range

    Set rangeToUpdate = GetRangeFromTable(cstGamesSheet, cstGamesTable, "Game date", pID)
    rangeToUpdate.Value = CStr(pDate)

End Sub

Public Sub GT_InsertGame(pWhiteBotName As String, pBlackBotName As String, pWinner As EState)
Dim ID As Integer

    ID = Log.CreateNewEntryAndGetID(cstGamesSheet, cstGamesTable)
    
    Call Log.GT_UpdateWhitePlayer(ID, pWhiteBotName)
    Call Log.GT_UpdateBlackPlayer(ID, pBlackBotName)
    Call Log.GT_UpdateGameDate(ID, Now)
    
    If pWinner = EState.BlackWin Or pWinner = EState.WhiteFailed Then
        Call Log.GT_UpdateWinner(ID, pBlackBotName)
    ElseIf pWinner = EState.WhiteWin Or pWinner = EState.BlackFailed Then
        Call Log.GT_UpdateWinner(ID, pWhiteBotName)
    ElseIf pWinner = EState.Draw Then
        Call Log.GT_UpdateWinner(ID, EnumString(EState.Draw))
    End If
    
    Call TT_InsertTurns(ID)
   
End Sub

Public Sub CG_ClearData()

    With Worksheets(cstCurrentGameSheet).ListObjects(cstCurrentTurnsTable)
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Rows.Delete
        End If
    End With
    
End Sub

Public Sub TT_ClearData()

    With Worksheets(cstTurnsSheet).ListObjects(cstTurnsTable)
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Rows.Delete
        End If
    End With
    
End Sub

Public Sub GT_ClearData()
    With Worksheets(cstGamesSheet).ListObjects(cstGamesTable)
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Rows.Delete
        End If
    End With
    
End Sub

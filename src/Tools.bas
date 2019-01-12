Attribute VB_Name = "Tools"
Option Explicit
Option Base 0

'/// https://www.exceltrick.com/formulas_macros/vba-wait-and-sleep-functions/?
#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If



'/// FONCTION   : Retourne un tableau contenant tous les pions de la couleur passée en PARAMÈTRE
'/// PARAMÈTRE  : EColor
'/// RETOUR     : Variant (tableau de pion)
Public Function GetPawns(ByVal pColor As EColor) As Variant
Dim pawnlist() As PawnModel
Dim pawnCandidate As PawnModel
Dim pawnCounter As Integer
Dim cell As Range

    pawnCounter = 0
    
    'pour chaque cellule du damier
    For Each cell In Range("Game")
        
        'on instancie un objet pion et on le construit avec la range de la cellule actuelle
        Set pawnCandidate = New PawnModel
        Call pawnCandidate.Build(cell)
        
        'si le pion est véritablement un pion et est de la couleur passée en PARAMÈTRE
        If pawnCandidate.IsPawn And pawnCandidate.Color = pColor Then
            
            'on redimensionne notre tableau
            ReDim Preserve pawnlist(pawnCounter)
            'on ajoute le pion au tableau de pion
            Set pawnlist(pawnCounter) = pawnCandidate
            'on incrzmente notre compteur
            pawnCounter = pawnCounter + 1
            
        End If
        
    Next cell
    
    'on associe le tableau de pion ainsi constitué au retour de la fonction
    GetPawns = pawnlist
    
End Function



'/// FONCTION   : Retourne vrai si le tableau est null ou vide
'/// PARAMÈTRE  : Variant
'/// RETOUR     : Booléen
Public Function IsArrayNullOrEmpty(ByVal pArray As Variant) As Boolean
    
    On Error Resume Next
    
    IsArrayNullOrEmpty = True
    
    If IsArray(pArray) And _
    Not IsError(LBound(pArray, 1)) And _
    LBound(pArray, 1) <= UBound(pArray, 1) And _
    Not IsEmpty(pArray) And _
    UBound(pArray) > -1 Then
        IsArrayNullOrEmpty = False
    End If
    
    Exit Function
    
End Function



'/// FONCTION   : Retourne vrai si l'élément "xxxToBeFound" est présent dans le tableau passé en paramètre
'/// PARAMÈTRE  : Element à trouver (ici l'élement est de type string), Variant (où chercher l'élément)
'/// RETOUR     : Booléen
Public Function IsInArray(ByVal pStringToBeFound As String, ByVal pArray As Variant) As Boolean
  IsInArray = (UBound(Filter(pArray, pStringToBeFound)) > -1)
End Function



'/// FONCTION   : Créer une string représentant le plateau de jeu
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : String
Public Function MakeBlueprintFromBoard() As String
Dim blueprint As String
Dim line As String
Dim mark As String
Dim boardRow As Range
Dim cell As Range
    
    For Each boardRow In Range("Game").Rows
        blueprint = blueprint + "|"
        For Each cell In boardRow.Cells
            With cell
                If (.Column + .Row) Mod 2 = 0 Then
                    mark = "-|"
                ElseIf .Value = "O" And .Font.Color = RGB(255, 255, 255) Then
                    mark = "w|"
                ElseIf .Value = Chr(169) And .Font.Color = RGB(255, 255, 255) Then
                    mark = "W|"
                ElseIf .Value = "O" And .Font.Color = RGB(0, 0, 0) Then
                    mark = "b|"
                ElseIf .Value = Chr(169) And .Font.Color = RGB(0, 0, 0) Then
                    mark = "B|"
                Else
                    mark = " |"
                End If
            End With
            blueprint = blueprint + mark
        Next cell
        blueprint = blueprint + vbNewLine
    Next boardRow
    
    blueprint = Left(blueprint, Len(blueprint) - 2)
    
    MakeBlueprintFromBoard = blueprint
End Function



'/// FONCTION   : Imprime une string représentant le plateau de jeu, sur le plateau
'/// PARAMÈTRE  : String
'/// RETOUR     : Aucun
Public Sub Compute(ByVal pBoardPattern As String)
Dim patternRows As Variant
Dim patternRow As Variant
Dim cellsMock As Variant
Dim cellMock As Variant
Dim rowCounter As Integer
Dim columnCounter
    
    rowCounter = 1
    patternRows = Split(pBoardPattern, vbNewLine)
    For Each patternRow In patternRows
        rowCounter = rowCounter + 1
        columnCounter = 1
        cellsMock = Split(Right(Left(patternRow, Len(patternRow) - 1), Len(patternRow) - 2), "|")
        For Each cellMock In cellsMock
            columnCounter = columnCounter + 1
            With Cells(rowCounter, columnCounter)
                Select Case cellMock
                    Case "w"
                        .Font.Color = RGB(255, 255, 255)
                        .Value = "O"
                    Case "b"
                        .Font.Color = RGB(0, 0, 0)
                        .Value = "O"
                    Case "W"
                        .Font.Color = RGB(255, 255, 255)
                        .Value = Chr(169)
                    Case "B"
                        .Font.Color = RGB(0, 0, 0)
                        .Value = Chr(169)
                    Case Else
                        .ClearContents
                End Select
            End With
        Next cellMock
    Next patternRow
End Sub

Public Sub RefreshScreen(Optional ByVal milliseconds As Integer)

    Application.ScreenUpdating = True
    DoEvents
    If Not IsMissing(milliseconds) Then Sleep milliseconds
    Application.ScreenUpdating = False
    
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoardModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'/// CONSTRUCTEUR: méthode appelé lors de l'instanciation de l'objet
'/// PARAMÈTRE   : Aucun
'/// RETOUR      : Aucun
Private Sub class_Initialize()
    If Not Range("K2").Value = "Restart" Then Call SetNamedRange
End Sub



'/// DESTRUCTEUR: méthode appelé lors de la destruction de l'objet
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Private Sub class_Terminate()
End Sub



'/// PROCÉDURE  : Initialise le plateau de jeu en positionnant les pions sur leur valeur par défaut
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub Initialisation()
Dim blueprint As String

    '             a b c d e f g h
    blueprint = "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "|b|-|b|-|b|-|b|-|" + vbNewLine + _
                "|-|b|-|b|-|b|-|b|" + vbNewLine + _
                "| |-| |-| |-| |-|" + vbNewLine + _
                "|-| |-| |-| |-| |" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|" + vbNewLine + _
                "|-|w|-|w|-|w|-|w|" + vbNewLine + _
                "|w|-|w|-|w|-|w|-|"
                
    Call Tools.Compute(blueprint)
End Sub



'/// PROCÉDURE  : formate la feuille Excel en plateau de jeu par défaut
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub FormatBoard()
Dim cell As Range

    Worksheets(1).Name = "BOARD"
    Worksheets("BOARD").Select

    'On met toutes les cases, du plateau et autour, vide et gris clair
    With Range("A1:N11")
        .ClearContents
        .Interior.Color = RGB(230, 230, 230)
        .MergeCells = False
    End With

    'on récupère toutes les cellules qui composent notre damier
    For Each cell In Range("B2:I9")
        'si la somme colonne + ligne d'une cellule est paire, alors on set la couleur claire, sinon la sombre
        If (cell.Column + cell.Row) Mod 2 = 0 Then
            'colore la cellule en blanc cassé
            cell.Interior.Color = RGB(255, 255, 200)
        Else
            'colore la cellule en marron
            cell.Interior.Color = RGB(150, 100, 50)
        End If
    Next cell
    
    Range("B1").Value = "A"
    Range("C1").Value = "B"
    Range("D1").Value = "C"
    Range("E1").Value = "D"
    Range("F1").Value = "E"
    Range("G1").Value = "F"
    Range("H1").Value = "G"
    Range("I1").Value = "H"
    
    Range("A2").Value = "1"
    Range("A3").Value = "2"
    Range("A4").Value = "3"
    Range("A5").Value = "4"
    Range("A6").Value = "5"
    Range("A7").Value = "6"
    Range("A8").Value = "7"
    Range("A9").Value = "8"
    
    
    'On modifie les paramètres de toutes les cellules comprises sur l'aire de jeu et a coté
    With Range("A1:N11")
        'Hauteur de la cellule
        .RowHeight = 25
        'Largeur de colonne
        .ColumnWidth = 4
        'Aligne le texte au centre de la cellule horizontalement
        .HorizontalAlignment = xlCenter
        'Aligne le texte au centre de la cellule verticalement
        .VerticalAlignment = xlCenter
         With .Font
            'Taille de la police
            .Size = 14
            'Met en gras
            .Bold = True
        End With
    End With
    
    'Fusionne les cellules
    Range("K2:M2").MergeCells = True
    'Bouton Restart
    With Range("K2:M2")
        'écris la valeur de la cellule
        .Value = "Restart"
        'Définit les contours par un trait continue
        .Borders.LineStyle = xlContinuous
        'Définit l'épaisseur du contour : gros
        .Borders.Weight = xlThick
        'Couleur du fond de cellule blanc
        .Interior.ColorIndex = 2
    End With
    
    'En-Tête de l'indicateur du tour en cours
    Range("K4:M4").MergeCells = True
    Range("K4:M4").Value = "Turn"
    
    'Indicateur couleur du tour en cours
    Range("K5:M5").MergeCells = True
    With Range("K5:M5")
        .Value = "White"
        .Interior.ColorIndex = 2
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    'Configuration du nombre de joueur
    Range("K7:M7").MergeCells = True
    With Range("K7:M7")
        .Value = "SinglePlayer"
        .Interior.ColorIndex = 2
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    'Pion en mémoire
    Range("K9:M9").MergeCells = True
    With Range("K9:M9")
        .Value = ""
        .Interior.ColorIndex = 2
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    'on masque toutes les colonnes et lignes au-delà en bas et à droite de la case Q14
    Range(Range("O12"), Range("O12").End(xlToRight)).EntireColumn.Hidden = True
    Range(Range("O12"), Range("O12").End(xlDown)).EntireRow.Hidden = True

    Call SetNamedRange

End Sub



'/// PROPRIÉTÉ  : Retourne le nom de la section dans laquelle se trouve la Range passée en paramètre
'/// PARAMÈTRE  : Range
'/// RETOUR     : ESection
Public Property Get GetSection(ByVal pCell As Range) As ESection

    'par défaut, on définit la valeur de retour à Hors limite
    GetSection = ESection.OutOfLimit
    
    'si la cellule est comprise danse une des sections connues du plateau, alors on renvoi la valeur de la section
    If Not (Application.Intersect(pCell, Range("Game")) Is Nothing) Then
        GetSection = ESection.Game
    ElseIf Not (Application.Intersect(pCell, Range("ConfigPlayer")) Is Nothing) Then
        GetSection = ESection.ConfigPlayer
    ElseIf Not (Application.Intersect(pCell, Range("Restart")) Is Nothing) Then
        GetSection = ESection.Restart
    End If

End Property



'/// PROPRIÉTÉ  : Retourne la valeur de l'énumération EColor associée de la valeur écrite sur le damier
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : EColor
Public Property Get TurnColor() As EColor

    If Range("TurnValue") = EnumString(EColor.White) Then
        TurnColor = EColor.White
    ElseIf Range("TurnValue") = EnumString(EColor.Black) Then
        TurnColor = EColor.Black
    Else
        'Si aucune valeur attendue n'est trouvée, on la paramètre à "White" par défaut
        TurnColor = EColor.White
        Range("TurnValue") = EnumString(EColor.White)
    End If
    
End Property



'/// PROPRIÉTÉ  : Ecrit la couleur du joueur du tour
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Property Let TurnColor(ByVal pColor As EColor)
    Range("TurnValue") = EnumString(pColor)
End Property



'/// PROPRIÉTÉ  : Ecrit l'etat de la partie
'/// PARAMÈTRE  : EState
'/// RETOUR     : Aucun
Public Property Let GameState(pState As EState)
    Select Case pState
        Case EState.WhiteWin
            Range("TurnValue") = "White Win"
        Case EState.BlackWin
            Range("TurnValue") = "Black Win"
        Case EState.Draw
            Range("TurnValue") = "Draw"
        Case EState.WhiteFailed
            Range("TurnValue") = "White Failed"
        Case EState.BlackFailed
            Range("TurnValue") = "Black Failed"
    End Select
End Property
    


'/// PROPRIÉTÉ  : Retourne la valeur de l'énumération EState associée a l'état de la partie
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : EState
Public Property Get GameState() As EState
    Select Case Range("TurnValue").Value
        Case "White Win"
            GameState = EState.WhiteWin
        Case "Black Win"
            GameState = EState.BlackWin
        Case "Draw"
            GameState = EState.Draw
        Case "White Failed"
            GameState = EState.WhiteFailed
        Case "Black Failed"
            GameState = EState.BlackFailed
        Case "White", "Black"
            GameState = EState.TheGameMustGoOn
        Case Else
            GameState = EState.InvalidState
    End Select
End Property



'/// PROPRIÉTÉ  : Retourne la configuration actuelle du plateau jeu (1 joueur, 2 joueurs, ia vs ia)
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : EConfig
Public Property Get ConfigPlayer() As EConfig
    
    If Range("ConfigPlayerValue") = EnumString(EConfig.SinglePlayer) Then
        ConfigPlayer = EConfig.SinglePlayer
    ElseIf Range("ConfigPlayerValue").Value = EnumString(EConfig.TwoPlayers) Then
        ConfigPlayer = EConfig.TwoPlayers
    ElseIf Range("ConfigPlayerValue").Value = EnumString(EConfig.Automate) Then
        ConfigPlayer = EConfig.Automate
    Else
        'Si aucune valeur attendue n'est trouvée, on la paramètre a "1 Player" par défaut
        ConfigPlayer = EConfig.SinglePlayer
        Range("ConfigPlayerValue") = EnumString(EConfig.SinglePlayer)
    End If
    
End Property



'/// PROPRIÉTÉ  : Modifie la configuration du plateau de jeu
'/// PARAMÈTRE  : EConfig
'/// RETOUR     : Aucun
Public Property Let ConfigPlayer(ByVal pConfig As EConfig)
        Range("ConfigPlayer") = EnumString(pConfig)
End Property



'/// PROPRIÉTÉ  : écris sur le plateau de jeu dans la zone "mémoire" la position du pion passé en paramètre
'/// PARAMÈTRE  : PawnModel
'/// RETOUR     : Aucun
Public Property Let Memory(ByVal pPawn As PawnModel)
Dim Column As String
Dim Row As String
Dim offset As Integer

    'permet d'écrire sur le plateau les coordonnées dans le référentiel du plateau et non le référentiel d’Excel
    'le damier commençant en "B2", on a un offset de 1 sur la colonne et de 1 sur la ligne par rapport à "A1"
    offset = 1
    
    'afin de récupérer la lettre correspondant à la colonne, on convertie la valeur de la colonne en caractère ASCII
    Column = Chr(64 + (pPawn.CurrentRange.Column - offset))
    'on convertie notre valeur numérique en chaine de caractère : Integer to String
    Row = CStr(pPawn.CurrentRange.Row - offset)
    
    Range("MemoryValue").Value = Column + Row
        
End Property



'/// PROPRIÉTÉ  : Retourne le pion écris en mémoire sur le plateau de jeu
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : PawnModel
Public Property Get Memory() As PawnModel
Dim offsetPawnPosition As String
Dim offsetPawnRange As Range
Dim pawnRange As Range
Dim offset As Integer
    
    Set Memory = New PawnModel
    
    offset = 1
    
    'si la range "mémoire" n'est pas vie
    If Not IsEmpty(Range("MemoryValue")) Then
        
        'on récupère la valeur écrite sur le plateau
        offsetPawnPosition = Range("MemoryValue").Value
        
        'récupère la cellule dans le référentiel du damier
        Set offsetPawnRange = Range(offsetPawnPosition)
        
        'récupère la véritable cellule dans le référentiel d’Excel
        Set pawnRange = Cells(offsetPawnRange.Row + offset, offsetPawnRange.Column + offset)
        
        'on construit le pion avec la cellule précédemment récupéré
        Call Memory.Build(pawnRange)
        
    End If
    
End Property



'/// PROCÉDURE  : Termine le tour de jeu en mettant a jour l'etat de la partie
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub EndTurn()
Dim blackPawns As Variant
Dim whitePawns As Variant
Dim enemyPawns As Variant
Dim pawn As Variant
Dim canEnemyMoves As Boolean
    
    ' search winner
    If Me.GameState = EState.TheGameMustGoOn Then
        blackPawns = Tools.GetPawns(EColor.Black)
        whitePawns = Tools.GetPawns(EColor.White)
        canEnemyMoves = False
        
        If Tools.IsArrayNullOrEmpty(blackPawns) Then
            Me.GameState = EState.WhiteWin
        ElseIf Tools.IsArrayNullOrEmpty(whitePawns) Then
            Me.GameState = EState.BlackWin
        Else
        
            If Me.TurnColor = EColor.White Then
                enemyPawns = blackPawns
            ElseIf Me.TurnColor = EColor.Black Then
                enemyPawns = whitePawns
            End If
            
            For Each pawn In enemyPawns
                If pawn.CanMove Or pawn.CanAttack Then
                    canEnemyMoves = True
                End If
            Next pawn
            
            ' si l'enemi peut attaquer on change la couleur pour le tour suivant
            If canEnemyMoves Then
                If Me.GameState = EState.TheGameMustGoOn Then
                    If Me.TurnColor = EColor.Black Then
                        Me.TurnColor = EColor.White
                    ElseIf Me.TurnColor = EColor.White Then
                        Me.TurnColor = EColor.Black
                    End If
                End If
            ' si aucun pion enemis ne peut se déplacer il y a égalité
            Else
                Me.GameState = EState.Draw
            End If
        End If
    End If
    
End Sub



'/// PROCÉDURE  : Vérifie la présence des plages nommées utilisées par le damier dans le classeur excel
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub SetNamedRange()

    'Accès par l'interface graphique : Formules -> Noms définis -> Gestionnaire de noms

    'On paramètre plusieurs emplacements par défaut
    ' Range("Game") = Range("B2:I9")
    ' Range("Restart") = Range("K2:M2")
    ' Range("Turn") = Range("K5:M5")
    ' Range("ConfigPlayer") = Range("K7:M7")
    ' Range("Memory") = Range("K9:M9")
    
    On Error Resume Next
    
    'Correspond à l'air de jeu du damier
    If IsError(Range("Game").Select) Then
        ActiveWorkbook.Names.Add Name:="Game", RefersToR1C1:="=Feuil1!R2C2:R9C9"
    End If
    
    'Correspond au bouton "Restart" à côté du damier
    If IsError(Range("Restart").Select) Then
        ActiveWorkbook.Names.Add Name:="Restart", RefersToR1C1:="=Feuil1!R2C11:R2C13"
    End If
    

    'Correspond à la zone indiquant la couleur du joueur du tour en cours
    If IsError(Range("Turn").Select) Then
        ActiveWorkbook.Names.Add Name:="Turn", RefersToR1C1:="=Feuil1!R5C11:R5C13"
    End If

    'afin de vérifier des égalités de valeur, il est nécessaire d'avoir une range contenant uniquement une seule cellule
    If IsError(Range("TurnValue").Select) Then
        ActiveWorkbook.Names.Add Name:="TurnValue", RefersToR1C1:="=Feuil1!R5C11"
    End If

    'Correspond au bouton de configuration du nombre de joueur
    If IsError(Range("ConfigPlayer").Select) Then
        ActiveWorkbook.Names.Add Name:="ConfigPlayer", RefersToR1C1:="=Feuil1!R7C11:R7C13"
    End If

    If IsError(Range("ConfigPlayerValue").Select) Then
        ActiveWorkbook.Names.Add Name:="ConfigPlayerValue", RefersToR1C1:="=Feuil1!R7C11"
    End If
    

    'Correspond à la zone tampon permettant de savoir si un pion est en mémoire
    If IsError(Range("Memory").Select) Then
        ActiveWorkbook.Names.Add Name:="Memory", RefersToR1C1:="=Feuil1!R9C11:R9C13"
    End If

    If IsError(Range("MemoryValue").Select) Then
        ActiveWorkbook.Names.Add Name:="MemoryValue", RefersToR1C1:="=Feuil1!R9C11"
    End If

End Sub



'/// PROPRIÉTÉ  : Retourne true si la partie est terminé
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Public Property Get IsGameOver() As Boolean

    IsGameOver = False
    
    If Not Me.GameState = EState.TheGameMustGoOn Then
        IsGameOver = True
    End If
    
End Property



'/// PROPRIÉTÉ  : Retourne un tableau comprenant les coefficients associés à la direction du déplacement
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Variant (tableau de la forme (x,y))
Public Property Get CardinalCoef(ByVal pDirection As EWindRose) As Variant
    
    If pDirection = EWindRose.SouthEast Then
        CardinalCoef = Array(1, 1)
    ElseIf pDirection = EWindRose.NorthEast Then
        CardinalCoef = Array(-1, 1)
    ElseIf pDirection = EWindRose.SouthWest Then
        CardinalCoef = Array(1, -1)
    ElseIf pDirection = EWindRose.NorthWest Then
        CardinalCoef = Array(-1, -1)
    End If
    
End Property

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

'/// CONSTRUCTEUR: m�thode appel� lors de l'instanciation de l'objet
'/// PARAM�TRE   : Aucun
'/// RETOUR      : Aucun
Private Sub class_Initialize()
    If Not Range("K2").Value = "Restart" Then Call SetNamedRange
End Sub



'/// DESTRUCTEUR: m�thode appel� lors de la destruction de l'objet
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Private Sub class_Terminate()
End Sub



'/// PROC�DURE  : Initialise le plateau de jeu en positionnant les pions sur leur valeur par d�faut
'/// PARAM�TRE  : Aucun
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



'/// PROC�DURE  : formate la feuille Excel en plateau de jeu par d�faut
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Public Sub FormatBoard()
Dim cell As Range
    
    'On met toutes les cases, du plateau et autour, vide et gris clair
    With Range("A1:N11")
        .ClearContents
        .Interior.Color = RGB(230, 230, 230)
        .MergeCells = False
    End With

    'on r�cup�re toutes les cellules qui composent notre damier
    For Each cell In Range("B2:I9")
        'si la somme colonne + ligne d'une cellule est paire, alors on set la couleur claire, sinon la sombre
        If (cell.Column + cell.Row) Mod 2 = 0 Then
            'colore la cellule en blanc cass�
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
    
    
    'On modifie les param�tres de toutes les cellules comprises sur l'aire de jeu et a cot�
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
        '�cris la valeur de la cellule
        .Value = "Restart"
        'D�finit les contours par un trait continue
        .Borders.LineStyle = xlContinuous
        'D�finit l'�paisseur du contour : gros
        .Borders.Weight = xlThick
        'Couleur du fond de cellule blanc
        .Interior.ColorIndex = 2
    End With
    
    'En-T�te de l'indicateur du tour en cours
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
    
    'Pion en m�moire
    Range("K9:M9").MergeCells = True
    With Range("K9:M9")
        .Value = ""
        .Interior.ColorIndex = 2
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    'on masque toutes les colonnes et lignes au-del� en bas et � droite de la case Q14
    Range(Range("O12"), Range("O12").End(xlToRight)).EntireColumn.Hidden = True
    Range(Range("O12"), Range("O12").End(xlDown)).EntireRow.Hidden = True

    Call SetNamedRange
    
End Sub



'/// PROPRI�T�  : Retourne le nom de la section dans laquelle se trouve la Range pass�e en param�tre
'/// PARAM�TRE  : Range
'/// RETOUR     : ESection
Public Property Get GetSection(ByVal pCell As Range) As ESection

    'par d�faut, on d�finit la valeur de retour � Hors limite
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



'/// PROPRI�T�  : Retourne la valeur de l'�num�ration EColor associ�e de la valeur �crite sur le damier
'/// PARAM�TRE  : Aucun
'/// RETOUR     : EColor
Public Property Get TurnColor() As EColor

    If Range("TurnValue") = EnumString(EColor.White) Then
        TurnColor = EColor.White
    ElseIf Range("TurnValue") = EnumString(EColor.Black) Then
        TurnColor = EColor.Black
    Else
        'Si aucune valeur attendue n'est trouv�e, on la param�tre � "White" par d�faut
        TurnColor = EColor.White
        Range("TurnValue") = EnumString(EColor.White)
    End If
    
End Property



'/// PROC�DURE  : Change la couleur affich�e sur le plateau de jeu
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Public Sub SwitchTurn()

    If Range("TurnValue") = EnumString(EColor.Black) Or Range("TurnValue") = EnumString(EColor.White) Then
        If Me.TurnColor = EColor.Black Then
            Me.TurnColor = EColor.White
        ElseIf Me.TurnColor = EColor.White Then
            Me.TurnColor = EColor.Black
        End If
    End If
    
End Sub



'/// PROPRI�T�  : Retourne la couleur du joueur du tour en cours
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Property Let TurnColor(ByVal pColor As EColor)
    Range("TurnValue") = EnumString(pColor)
End Property



'/// PROPRI�T�  : Retourne la configuration actuelle du plateau jeu (1 joueur, 2 joueurs, ia vs ia)
'/// PARAM�TRE  : Aucun
'/// RETOUR     : EConfig
Public Property Get ConfigPlayer() As EConfig
    
    If Range("ConfigPlayerValue") = EnumString(EConfig.SinglePlayer) Then
        ConfigPlayer = EConfig.SinglePlayer
    ElseIf Range("ConfigPlayerValue").Value = EnumString(EConfig.TwoPlayers) Then
        ConfigPlayer = EConfig.TwoPlayers
    ElseIf Range("ConfigPlayerValue").Value = EnumString(EConfig.Automate) Then
        ConfigPlayer = EConfig.Automate
    Else
        'Si aucune valeur attendue n'est trouv�e, on la param�tre a "1 Player" par d�faut
        ConfigPlayer = EConfig.SinglePlayer
        Range("ConfigPlayerValue") = EnumString(EConfig.SinglePlayer)
    End If
    
End Property



'/// PROPRI�T�  : Modifie la configuration du plateau de jeu
'/// PARAM�TRE  : EConfig
'/// RETOUR     : Aucun
Public Property Let ConfigPlayer(ByVal pConfig As EConfig)
        Range("ConfigPlayer") = EnumString(pConfig)
End Property



'/// PROPRI�T�  : �cris sur le plateau de jeu dans la zone "m�moire" la position du pion pass� en param�tre
'/// PARAM�TRE  : PawnModel
'/// RETOUR     : Aucun
Public Property Let Memory(ByVal pPawn As PawnModel)
Dim Column As String
Dim Row As String
Dim offset As Integer

    'permet d'�crire sur le plateau les coordonn�es dans le r�f�rentiel du plateau et non le r�f�rentiel d�Excel
    'le damier commen�ant en "B2", on a un offset de 1 sur la colonne et de 1 sur la ligne par rapport � "A1"
    offset = 1
    
    'afin de r�cup�rer la lettre correspondant � la colonne, on convertie la valeur de la colonne en caract�re ASCII
    Column = Chr(64 + (pPawn.CurrentRange.Column - offset))
    'on convertie notre valeur num�rique en chaine de caract�re : Integer to String
    Row = CStr(pPawn.CurrentRange.Row - offset)
    
    Range("MemoryValue").Value = Column + Row
        
End Property



'/// PROPRI�T�  : Retourne le pion �cris en m�moire sur le plateau de jeu
'/// PARAM�TRE  : Aucun
'/// RETOUR     : PawnModel
Public Property Get Memory() As PawnModel
Dim offsetPawnPosition As String
Dim offsetPawnRange As Range
Dim pawnRange As Range
Dim offset As Integer
    
    Set Memory = New PawnModel
    
    offset = 1
    
    'si la range "m�moire" n'est pas vie
    If Not IsEmpty(Range("MemoryValue")) Then
        
        'on r�cup�re la valeur �crite sur le plateau
        offsetPawnPosition = Range("MemoryValue").Value
        
        'r�cup�re la cellule dans le r�f�rentiel du damier
        Set offsetPawnRange = Range(offsetPawnPosition)
        
        'r�cup�re la v�ritable cellule dans le r�f�rentiel d�Excel
        Set pawnRange = Cells(offsetPawnRange.Row + offset, offsetPawnRange.Column + offset)
        
        'on construit le pion avec la cellule pr�c�demment r�cup�r�
        Call Memory.Build(pawnRange)
        
    End If
    
End Property



'/// PROC�DURE  : Ecris sur le plateau de jeu "You Win" s�il ne reste qu'un seul joueur ou "Draw" en cas d'�galit�
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Public Sub SearchWinner()
Dim blackPawns As Variant
Dim whitePawns As Variant
Dim enemyPawns As Variant
Dim pawn As Variant
Dim canEnemyAttack As Boolean
    
    If Range("TurnValue") = "White" Or Range("TurnValue") = "Black" Then
        blackPawns = Tools.GetPawns(EColor.Black)
        whitePawns = Tools.GetPawns(EColor.White)
        
        If Tools.IsArrayNullOrEmpty(blackPawns) Or _
           Tools.IsArrayNullOrEmpty(whitePawns) Then
        
            'si une couleur n'a plus de pion
            Range("TurnValue") = "You Win"
            
        Else
        
            If Me.TurnColor = EColor.White Then
                enemyPawns = blackPawns
            ElseIf Me.TurnColor = EColor.Black Then
                enemyPawns = whitePawns
            End If
            
            For Each pawn In enemyPawns
                If pawn.CanMove Or pawn.CanAttack Then
                    canEnemyAttack = True
                End If
            Next pawn
            
            If Not canEnemyAttack Then
                Range("TurnValue") = "Draw"
            End If
        End If
    End If

End Sub


'/// PROC�DURE  : V�rifie la pr�sence des plages nomm�es, utilis�s par le damier, dans le classeur excel
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Public Sub SetNamedRange()

    'Acc�s par l'interface graphique : Formules -> Noms d�finis -> Gestionnaire de noms

    'On instancie l'objet Plateau en param�trant plusieurs emplacements par d�faut
    
    On Error Resume Next
    
    'Correspond � l'air de jeu du damier
    If IsError(Range("Game").Select) Then
        ActiveWorkbook.Names.Add Name:="Game", RefersToR1C1:="=Feuil1!R2C2:R9C9"
    End If
    
    'Correspond au bouton "Restart" � c�t� du damier
    If IsError(Range("Restart").Select) Then
        ActiveWorkbook.Names.Add Name:="Restart", RefersToR1C1:="=Feuil1!R2C11:R2C13"
    End If
    

    'Correspond � la zone indiquant la couleur du joueur du tour en cours
    If IsError(Range("Turn").Select) Then
        ActiveWorkbook.Names.Add Name:="Turn", RefersToR1C1:="=Feuil1!R5C11:R5C13"
    End If

    'afin de v�rifier des �galit�s de valeur, il est n�cessaire d'avoir une range contenant uniquement une seule cellule
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
    

    'Correspond � la zone tampon permettant de savoir si un pion est en m�moire
    If IsError(Range("Memory").Select) Then
        ActiveWorkbook.Names.Add Name:="Memory", RefersToR1C1:="=Feuil1!R9C11:R9C13"
    End If

    If IsError(Range("MemoryValue").Select) Then
        ActiveWorkbook.Names.Add Name:="MemoryValue", RefersToR1C1:="=Feuil1!R9C11"
    End If

End Sub



'/// PROPRI�T�  : Retourne true si "You win" est �cris sur le plateau
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Boolean
Property Get IsWinner() As Boolean

    IsWinner = False
    
    If Range("TurnValue") <> EnumString(EColor.White) And _
       Range("TurnValue") <> EnumString(EColor.Black) Then
        IsWinner = True
    End If
    
End Property



'/// PROPRI�T�  : Retourne un tableau comprenant les coefficients associ�s � la direction du d�placement
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Variant (tableau de la forme (x,y))
Property Get CardinalCoef(ByVal pDirection As EWindRose) As Variant
    
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

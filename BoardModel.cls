VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoardModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'/// CONSTRUCTEUR: methode appelé lors de l'instanciation de l'objet
'/// PARAMÈTRE   : Aucun (evenement constructeur)
'/// RETOUR      : Aucun (evenement)
Private Sub class_Initialize()


    'On instancie l'objet Plateau en parametrant plusieurs emplacement par defaut
    
    On Error GoTo Catch
    
    'cette condition permet en réalité de tester si le paramétrage des range nommées a été fait (exemple Range("Game"), plutot que Range("B2:I9"))
    If Not (Application.Intersect(Range("B2"), Range("Game")) Is Nothing) Then
        'Alors les noms de plages sont déja définis dans le classeur
    End If
    
    Exit Sub
    
Catch:
    
    'Acces par l'interface graphique : Formules -> Noms définis -> Gestionnaire de noms
    
    'Correspond a l'air de jeu du damier
    ActiveWorkbook.Names.Add Name:="Game", RefersToR1C1:="=Feuil1!R2C2:R9C9"
    
    'Correspond au bouton "Restart" a coté du damier
    ActiveWorkbook.Names.Add Name:="Restart", RefersToR1C1:="=Feuil1!R2C11:R2C13"
    
    'Correspond à la zone indiquant la couleurs du joueur du tour en cours
    ActiveWorkbook.Names.Add Name:="Turn", RefersToR1C1:="=Feuil1!R5C11:R5C13"
    
    'Correspond au bouton de configuration du nombre de joueur
    ActiveWorkbook.Names.Add Name:="ConfigPlayer", RefersToR1C1:="=Feuil1!R7C11:R7C13"
    
    'Correspond a la zone tampon permettant de savoir si un pion est en mémoire
    ActiveWorkbook.Names.Add Name:="Memory", RefersToR1C1:="=Feuil1!R9C11:R9C13"
    
    
    'afin de vérifier des égalité de valeur, il est necessaire d'avoir une range contenant uniquement une seule cellule
    ActiveWorkbook.Names.Add Name:="TurnValue", RefersToR1C1:="=Feuil1!R5C11"
    ActiveWorkbook.Names.Add Name:="ConfigPlayerValue", RefersToR1C1:="=Feuil1!R7C11"
    ActiveWorkbook.Names.Add Name:="MemoryValue", RefersToR1C1:="=Feuil1!R9C11"
    
End Sub



'/// DESTRUCTEUR: methode appelé lors de la destruction de l'objet
'/// PARAMÈTRE  : Aucun (evenement destructeur)
'/// RETOUR     : Aucun (evenemnt)
Private Sub class_Terminate()

End Sub



'/// PROPRIÉTÉ  : Retourne le nom de la section dans laquelle se trouve la Range passée en parametre
'/// PARAMÈTRE  : Range
'/// RETOUR     : ESection
Public Property Get GetSection(cell As Range) As ESection

    'par defaut, on definie la valeur de retour a Hors limite
    GetSection = ESection.OutOfLimit
    
    'si la cellule est comprise danse une des sections connues du plateau, alors on renvoit la valeur de la section
    If Not (Application.Intersect(cell, Range("Game")) Is Nothing) Then
        GetSection = ESection.Game
    ElseIf Not (Application.Intersect(cell, Range("ConfigPlayer")) Is Nothing) Then
        GetSection = ESection.ConfigPlayer
    ElseIf Not (Application.Intersect(cell, Range("Restart")) Is Nothing) Then
        GetSection = ESection.Restart
    End If

End Property



'/// PROPRIÉTÉ  : Retourne la valeur de l'enumeration EColor associée de la valeur écrite sur le damier
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : EColor
Public Property Get CurrentTurn() As EColor

    If Range("TurnValue") = "White" Then
        CurrentTurn = EColor.White
    ElseIf Range("TurnValue") = "Black" Then
        CurrentTurn = EColor.Black
    Else
        'Si aucune valeur attendue n'est trouvé, on la parametre a "White" par defaut
        CurrentTurn = EColor.White
        Range("TurnValue") = "White"
    End If
    
End Property



'/// PROCÉDURE  : Change la couleur afficher sur le plateau de jeu
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub SwitchTurn()

    If Range("TurnValue") <> "You Win" Then
        If CurrentTurn = EColor.Black Then
            CurrentTurn = EColor.White
        ElseIf CurrentTurn = EColor.White Then
            CurrentTurn = EColor.Black
        End If
    End If
    
End Sub



'/// PROPRIÉTÉ  : Retourne
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Property Let CurrentTurn(pColor As EColor)

    If pColor = EColor.Black Then
        Range("TurnValue") = "Black"
    ElseIf pColor = EColor.White Then
        Range("TurnValue") = "White"
    End If
    
End Property



'/// PROPRIÉTÉ  : Retourne la valeur écrite sur le damier puis renvoie la valeur de l'enumeration EConfig associée
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : EConfig
Public Property Get ConfigPlayer() As EConfig
    
    If Range("ConfigPlayerValue") = "1 Players" Then
        ConfigPlayer = EConfig.SinglePlayer
    ElseIf Range("ConfigPlayerValue").Value = "2 Players" Then
        ConfigPlayer = EConfig.TwoPlayer
    ElseIf Range("ConfigPlayerValue").Value = "Automa" Then
        ConfigPlayer = EConfig.Automate
    Else
        'Si aucune valeur attendue n'est trouvé, on la parametre a "1 Player" par defaut
        ConfigPlayer = SinglePlayer
        Range("ConfigPlayerValue") = "1 Player"
    End If
    
End Property



'/// PROPRIÉTÉ  : Ecris la valeur passé en parametre sur la plateau de jeu
'/// PARAMÈTRE  : EConfig
'/// RETOUR     : Aucun
Public Property Let ConfigPlayer(Config As EConfig)

    If Config = EConfig.SinglePlayer Then
        Range("ConfigPlayer") = "1 Players"
    ElseIf Config = EConfig.TwoPlayer Then
        Range("ConfigPlayer") = "2 Players"
    ElseIf Config = EConfig.Automate Then
        Range("ConfigPlayer") = "Automa"
    End If
        
End Property



'/// PROPRIÉTÉ  : Écris sur le plateau de jeu dans la zone "mémoire" la position du pion passé en parametre
'/// PARAMÈTRE  : PawnModel
'/// RETOUR     : Acun
Public Property Let Memory(pawn As PawnModel)

    Dim Column As String
    Dim Row As String
    
    'permet d'écrire sur le plateau les coordonées dans le référentiel du plateau et non le référentiel d'excel
    Dim offset As Integer
    
    'le damier commencant en "B2", on a un offset de 1 sur la colonne et de 1 sur l ligne par rapport a "A1"
    offset = 1
    
    'afin de récuper la lettre correspondant a la colonne, on convertie la valeur de la colonne en charactere ASCII
    Column = Chr(64 + (pawn.CurrentRange.Column - offset))
    'on convertie notre valeur numéric en chaine de charactere : Integer to String
    Row = CStr(pawn.CurrentRange.Row - offset)
    
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
        
        'on récupere la valeur écrite sur le plateau
        offsetPawnPosition = Range("MemoryValue").Value
        
        'récupre la cellule dans le référentiel du damier
        Set offsetPawnRange = Range(offsetPawnPosition)
        
        'récupere la veritable cellule dans le référentiel d'excel
        Set pawnRange = Cells(offsetPawnRange.Row + offset, offsetPawnRange.Column + offset)
        
        'on construit le pion avec la cellule précédemment récupéré
        Call Memory.Build(pawnRange)
        
    End If
    
End Property



'/// PROPRIÉTÉ  : Ecris sur la plateau de jeu "You Win" si il ne reste qu'un seul joueur
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub SearchWinner()

    Dim whiteFound As Boolean, blackFound As Boolean
    Dim pawn As PawnModel
    
    whiteFound = False
    blackFound = False
    
    'pour chaque cellule du damier
    For Each cell In Range("Game")
        
        'on instancie et construie un pion avec la cellule actuelle
        Set pawn = New PawnModel
        Call pawn.Build(Cells(cell.Row, cell.Column))
        
        'si le pion est véritablement un pion
        If pawn.IsPawn Then
        
            'si le pion est noir
            If pawn.Color = EColor.Black Then
                blackFound = True
                
            'si le pion est blanc
            Else
                whiteFound = True
            End If
            
            'si on a trouvé un pion blanc et un pion noir
            If whiteFound And blackFound Then
                Exit For
            End If
            
        End If
    Next cell
    
    'si une des couleurs n'est pas trouvé
    If Not whiteFound Or Not blackFound Then
        Range("TurnValue") = "You Win"
    End If
    
End Sub



'/// PROPRIÉTÉ  : Retourne true si "You win" est écris sur le plateau
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Property Get IsWinner() As Boolean

    IsWinner = False
    
    If Range("TurnValue") = "You Win" Then
        IsWinner = True
    End If
    
End Property



'/// PROPRIÉTÉ  : Retourne un tableau comprenant les coefficients associés à la direction du déplacement
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Variant (tableau de la forme (x,y))
Property Get CardinalCoef(Direction As EWindRose) As Variant
    
    If Direction = EWindRose.SouthEast Then
        CardinalCoef = Array(1, 1)
    ElseIf Direction = EWindRose.NorthEast Then
        CardinalCoef = Array(-1, 1)
    ElseIf Direction = EWindRose.SouthWest Then
        CardinalCoef = Array(1, -1)
    ElseIf Direction = EWindRose.NorthWest Then
        CardinalCoef = Array(-1, -1)
    End If
    
End Property




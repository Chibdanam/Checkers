Attribute VB_Name = "BoardConstructeur"
Option Explicit
Option Base 0

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
        .Value = "1 Player"
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

    Call BoardConstructeur.SetNameRanged
    
End Sub



'/// PROC�DURE  : Initialise le plateau de jeu en positionnant les pions sur leur valeur par d�faut
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Public Sub Initialisation()
Dim cell As Range
    'supprime toutes les valeurs �crites sur le damier
    Range("B2:I9").ClearContents
    
    'boucle sur toutes les cellules comprises sur les 3 premi�res et sur les 3 derni�res lignes du damier
    For Each cell In Union(Range("B7:I9"), Range("B2:I4"))
        'Si la somme de l'index de colonne et de l'index de ligne est impaire, alors on positionne un pion
        If (cell.Column + cell.Row) Mod 2 <> 0 Then
            'Un pion est symbolis� par la lettre O majuscule
            cell.Value = "O"
        End If
    Next cell

    'associe la partie basse du damier a la couleur blanche
    Range("B7:I9").Font.Color = RGB(255, 255, 255)
    'noir
    Range("B2:I4").Font.Color = RGB(0, 0, 0)

End Sub



'/// PROC�DURE  : V�rifie la pr�sence des plages nomm�es, utilis�s par le damier, dans le classeur excel
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Public Sub SetNameRanged()

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

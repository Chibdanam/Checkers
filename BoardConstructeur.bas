Attribute VB_Name = "BoardConstructeur"
'/// PROCÉDURE  : formate la feuille excel en plateau de jeu par defaut
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun (sub)
Public Sub FormatBoard()

    'variable temporaire qui va nous permettre d'agir sur chaque cellule du damier une par une
    Dim cell As Range
    
    'On met toutes les cases, du plateau et autour, vide et gris clair
    With Range("A1:N11")
        .ClearContents
        .Interior.Color = RGB(230, 230, 230)
        .MergeCells = False
    End With

    'on recupere toute les cellules qui composent notre damier
    For Each cell In Range("B2:I9")
        'si la somme colonne + ligne d'une cellule est paire, alors on set la couleur clair, sinon la sombre
        If (cell.Column + cell.Row) Mod 2 = 0 Then
            'color la cellule en blanc cassé
            cell.Interior.Color = 13434879
        Else
            'color la cellule en marron
            cell.Interior.Color = 3368601
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
    
    
    'On modifie les parametres de toutes les cellules comprises sur l'air de jeu et a coté
    With Range("A1:N11")
        'Hauteur de la cellule
        .RowHeight = 25
        'Largeur de colonne
        .ColumnWidth = 4
        'Aligne le texte au centre de le cellule horizontalement
        .HorizontalAlignment = xlCenter
        'Aligne le texte au centre de le cellule verticalement
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
        'Écris la valeur de la cellule
        .Value = "Restart"
        'Définit les contours par un trait continue
        .Borders.LineStyle = xlContinuous
        'Définit l'épaisseur du contour : gros
        .Borders.Weight = xlThick
        'Couleur du fond de cellule blanc
        .Interior.ColorIndex = 2
    End With
    
    'En-Tete de l'indicateur du tour en cours
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
    
    'Pion en mémoire
    Range("K9:M9").MergeCells = True
    With Range("K9:M9")
        .Value = ""
        .Interior.ColorIndex = 2
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    
    'on masque toutes les colonnes et lignes au dela en bas et a droite de la case Q14
    Range(Range("O12"), Range("O12").End(xlToRight)).EntireColumn.Hidden = True
    Range(Range("O12"), Range("O12").End(xlDown)).EntireRow.Hidden = True
End Sub



'/// PROCÉDURE  : Initialise le plateau de jeu en positionnant les pions sur leur valeur par defaut
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun (sub)
Public Sub Initalisation()

    'supprime toutes les valeurs écrites sur le damier
    Range("B2:I9").ClearContents
    
    'boucle sur toutes les cellules comprises sur les 3 premiere et sur les 3 dernieres lignes du damier
    For Each cell In Union(Range("B7:I9"), Range("B2:I4"))
        'Si la somme de l'index de colonne et de l'index de ligne est impaire, alors on positionne un pion
        If (cell.Column + cell.Row) Mod 2 <> 0 Then
            'Un pion est symbolisé par la lettre O majuscule
            cell.Value = "O"
        End If
    Next cell

    'associe la partie basse du damier a la couleur blanche
    Range("B7:I9").Font.ColorIndex = 2
    'noir
    Range("B2:I4").Font.ColorIndex = -4105

End Sub







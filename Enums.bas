Attribute VB_Name = "Enums"
'une enum permet de d�finir une liste de contsantes (�num�ration).

'/// �NUM�RATION : Repr�sente les deux couleurs que peuvent avoir les pions
'/// VALEURS     : Blanc, Noir
Public Enum EColor
    White
    Black
End Enum

'/// �NUM�RATION : Repr�sente la configuration du jeu
'/// VALEURS     : 1 joueur, 2 joueurs, IA vs IA
Public Enum EConfig
    SinglePlayer
    TwoPlayer
    Automate
End Enum

'/// �NUM�RATION : Repr�sente une zone du plateau de jeu
'/// VALEURS     : Le jeu (le damier), Le bouton restaart, La zone affichant la configuration du jeu, Le reste
Public Enum ESection
    Game
    Restart
    ConfigPlayer
    OutOfLimit
End Enum

'/// �NUM�RATION : Repr�sente les 4 diagonales possible pour un d�placer
'/// VALEURS     : Nord Est, Nord Ouest, Sud Est, Sud Ouest
Public Enum EWindRose
    NorthEast
    NorthWest
    SouthEast
    SouthWest
End Enum

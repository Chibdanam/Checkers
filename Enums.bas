Attribute VB_Name = "Enums"
'une enum permet de d�finir un groupe de constantes li�es afin de cr�er une �num�ration.
'par convention, il convient de prefixer une enum par "E"

'/// �NUM�RATION : Repr�sente les deux couleurs que peuvent avoir les pions
'/// VALEURS     : Blanc, Noir
Enum EColor
    White
    Black
    WhiteMoveStep
    BlackMoveStep
End Enum

'/// �NUM�RATION : Repr�sente la configuration du jeu
'/// VALEURS     : 1 joueur, 2 joueurs, IA vs IA
Enum EConfig
    SinglePlayer
    TwoPlayer
    Automate
End Enum

'/// �NUM�RATION : Repr�sente une zone du plateau de jeu
'/// VALEURS     : Le jeu (le damier), Le bouton restaart, La zone affichant la configuration du jeu, Le reste
Enum ESection
    Game
    Restart
    ConfigPlayer
    OutOfLimit
End Enum

'/// �NUM�RATION : Repr�sente les 4 diagonales possible pour un d�placer
'/// VALEURS     : Nord Est, Nord Ouest, Sud Est, Sud Ouest
Enum EWindRose
    NorthEast
    NorthWest
    SouthEast
    SouthWest
End Enum


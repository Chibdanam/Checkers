Attribute VB_Name = "Enums"
'une enum permet de définir un groupe de constantes liées afin de créer une énumération.
'par convention, il convient de prefixer une enum par "E"

'/// ÉNUMÉRATION : Représente les deux couleurs que peuvent avoir les pions
'/// VALEURS     : Blanc, Noir
Enum EColor
    White
    Black
    WhiteMoveStep
    BlackMoveStep
End Enum

'/// ÉNUMÉRATION : Représente la configuration du jeu
'/// VALEURS     : 1 joueur, 2 joueurs, IA vs IA
Enum EConfig
    SinglePlayer
    TwoPlayer
    Automate
End Enum

'/// ÉNUMÉRATION : Représente une zone du plateau de jeu
'/// VALEURS     : Le jeu (le damier), Le bouton restaart, La zone affichant la configuration du jeu, Le reste
Enum ESection
    Game
    Restart
    ConfigPlayer
    OutOfLimit
End Enum

'/// ÉNUMÉRATION : Représente les 4 diagonales possible pour un déplacer
'/// VALEURS     : Nord Est, Nord Ouest, Sud Est, Sud Ouest
Enum EWindRose
    NorthEast
    NorthWest
    SouthEast
    SouthWest
End Enum


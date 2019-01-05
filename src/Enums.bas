Attribute VB_Name = "Enums"
'une enum permet de définir une liste de contsantes (énumération).

'/// ÉNUMÉRATION : Représente les deux couleurs que peuvent avoir les pions
'/// VALEURS     : Blanc, Noir
Public Enum EColor
    White = 0
    Black = 1
End Enum

'/// ÉNUMÉRATION : Représente la configuration du jeu
'/// VALEURS     : 1 joueur, 2 joueurs, IA vs IA
Public Enum EConfig
    SinglePlayer = 10
    TwoPlayers = 11
    Automate = 12
End Enum

'/// ÉNUMÉRATION : Représente une zone du plateau de jeu
'/// VALEURS     : Le jeu (le damier), Le bouton restaart, La zone affichant la configuration du jeu, Le reste
Public Enum ESection
    Game = 20
    Restart = 21
    ConfigPlayer = 22
    OutOfLimit = 23
End Enum

'/// ÉNUMÉRATION : Représente les 4 diagonales possible pour un déplacer
'/// VALEURS     : Nord Est, Nord Ouest, Sud Est, Sud Ouest
Public Enum EWindRose
    NorthEast = 30
    NorthWest = 31
    SouthEast = 32
    SouthWest = 33
End Enum

Public Enum EState
    TheGameMustGoOn = 40
    WhiteWin = 41
    BlackWin = 42
    Draw = 43
    WhiteFailed = 44
    BlackFailed = 45
    InvalidState = 46
End Enum



'/// FONCTION   : Retourne la valeur de l'énum sous la forme de string
'/// PARAMÈTRE  : Enum
'/// RETOUR     : String
Public Function EnumString(pEnum As Variant) As String
    Select Case pEnum

        'EColor
        Case EColor.White
            EnumString = "White"
        Case EColor.Black
            EnumString = "Black"

        'EConfig
        Case EConfig.SinglePlayer
            EnumString = "SinglePlayer"
        Case EConfig.TwoPlayers
            EnumString = "TwoPlayers"
        Case EConfig.Automate
            EnumString = "Automate"

        'ESection
        Case ESection.Game
            EnumString = "Game"
        Case ESection.Restart
            EnumString = "Restart"
        Case ESection.ConfigPlayer
            EnumString = "ConfigPlayer"
        Case ESection.OutOfLimit
            EnumString = "OutOfLimit"
            
            
        'EWindRose
        Case EWindRose.NorthEast
            EnumString = "NorthEast"
        Case EWindRose.NorthWest
            EnumString = "NorthWest"
        Case EWindRose.SouthEast
            EnumString = "SouthEast"
        Case EWindRose.SouthWest
            EnumString = "SouthWest"
            
            
        'EState
        Case EState.TheGameMustGoOn
            EnumString = "TheGameMustGoOn"
        Case EState.WhiteWin
            EnumString = "WhiteWin"
        Case EState.BlackWin
            EnumString = "BlackWin"
        Case EState.Draw
            EnumString = "Draw"
        Case EState.WhiteFailed
            EnumString = "WhiteFailed"
        Case EState.BlackFailed
            EnumString = "BlackFailed"
        Case EState.InvalidState
            EnumString = "InvalidState"

    End Select
End Function

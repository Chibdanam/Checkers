VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit                             '/// Toutes les variables doivent être déclarées
Option Base 0                               '/// Les Arrays comment à "0"



'/// PROCÉDURE  : Se déclenche lorsqu'un double clic survient sur la feuille et transmet en paramètre la range de la cellule cliquée
'/// PARAMÈTRE  : Range, Boolean
'/// RETOUR     : Aucun (évènement)
Private Sub Worksheet_BeforeDoubleClick(ByVal pTarget As Range, pCancel As Boolean)
Dim board As BoardModel
Dim turnCount As Integer
    
    Application.ScreenUpdating = False
    
    'Si Cancel a sa valeur par défaut ou Égal à False, alors après la fin de cette procédure,
    'Excel va continuer en traitant le clic droit comme il le fait par défaut.
    
    'Si Cancel est à True, alors Excel ne va pas faire son traitement habituel,
    'Excel considère que l’évènement du clic droit est entièrement pris en charge par la procédure.
    pCancel = True
    
    'on instancie un objet correspondant à notre plateau de jeu
    Set board = New BoardModel

    'selon la section dans laquelle on a cliqué
    Select Case board.GetSection(pTarget)
    
        'si on a cliqué sur le damier
        Case ESection.Game
        
            'selon la valeur actuelle de configuration de la partie
            Select Case board.ConfigPlayer
            
                'si en mode 2 joueurs
                Case EConfig.TwoPlayers
                    
                    'si l'action correspond au mouvement d'un pion
                    If Player.Run(pTarget) Then
                    
                        'vérifie s'il y a un vainqueur
                        Call board.SearchWinner

                        'change la couleur afficher sur le plateau
                        Call board.SwitchTurn
                    
                    End If
                    
                'si en mode 1 joueur
                Case EConfig.SinglePlayer
                
                    'si l'action correspond au mouvement d'un pion
                    If Player.Run(pTarget) Then
                        Call board.SearchWinner
                        Call board.SwitchTurn
                        
                        's'il n'y a pas de vainqueur
                        If Not board.IsWinner Then
                            
                            'on appel le bot
                            Call BotManager.Run(board.TurnColor)

                            'vérifie s'il y a un vainqueur
                            Call board.SearchWinner
                            
                            'change la couleur afficher sur le plateau
                            Call board.SwitchTurn
                                    
                        End If
                    End If
                
                'si en mode IA vs IA
                Case EConfig.Automate
                    
                    turnCount = 0
                    
                    'tant qu'il n'y a pas de vainqueur et qu'on a joué moins de 300 tours
                    While Not board.IsWinner And turnCount < 300
                    
                        Call BotManager.Run(board.TurnColor)
                        Call board.SearchWinner
                        Call board.SwitchTurn
                        
                        Call Tools.RefreshScreen(200)
                        
                        turnCount = turnCount + 1
                        
                    Wend
                    
                    
            End Select
                
        'cas de clic sur la configuration des joueurs, on incrémente la configuration
        Case ESection.ConfigPlayer
        
            'on incrémente le type de partie dans l'ordre -> 1joueurs, 2joueurs, ia vs ia
            If board.ConfigPlayer = EConfig.SinglePlayer Then
                board.ConfigPlayer = EConfig.TwoPlayers
            ElseIf board.ConfigPlayer = EConfig.TwoPlayers Then
                board.ConfigPlayer = EConfig.Automate
            Else
                board.ConfigPlayer = EConfig.SinglePlayer
            End If
        
        'cas de clic sur le bouton restart
        Case ESection.Restart
            'formate le plateau de jeux sur la feuille Excel
            Call board.FormatBoard
            'initialise les pions à leur position de départ
            Call board.Initialisation
            '
            Call Log.NewGame
    End Select
    Application.ScreenUpdating = True
End Sub

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
Dim pawn As PawnModel
Dim enemyPawn As PawnModel
Dim ysnp As YouShallNotPassModel
Dim deamonReapled As Integer
Dim turnCount As Integer

    'Si Cancel a sa valeur par défaut ou Égal à False, alors après la fin de cette procédure,
    'Excel va continuer en traitant le clic droit comme il le fait par défaut.
    
    'Si Cancel est à True, alors Excel ne va pas faire son traitement habituel,
    'Excel considère que l’évènement du clic droit est entièrement pris en charge par la procédure.
    pCancel = True
    
    'on instancie un objet correspondant à notre plateau de jeu
    Set board = New BoardModel
   
    'représente le compteur d'erreurs effectuées par le bot
    deamonReapled = 0

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

                        'change la couleur afficher sur le plateau
                        Call board.SwitchTurn
                    
                        'vérifie s'il y a un vainqueur
                        Call board.SearchWinner
                    
                    End If
                    
                'si en mode 1 joueur
                Case EConfig.SinglePlayer
                
                    'si l'action correspond au mouvement d'un pion
                    If Player.Run(pTarget) Then
                        Call board.SwitchTurn
                        Call board.SearchWinner
                        
                        's'il n'y a pas de vainqueur
                        If Not board.IsWinner Then

                            'instancie le vérificateur
                            Set ysnp = New YouShallNotPassModel

                            'sauvegarde la configuration du plateau avant le tour du bot
                            Call ysnp.Snapshot

                            Call Bot.Run

                        'si les actions du bot correspondent a un mouvement accépté
                        If ysnp.IsSuccess Then
                            deamonReapled = 0
                            Call board.SwitchTurn
                            Call board.SearchWinner
                        Else
                            deamonReapled = deamonReapled + 1
                            Call ysnp.RollBack()
                        End If
                        End If
                    End If
                
                'si en mode IA vs IA
                Case EConfig.Automate
                    
                    Set ysnp = New YouShallNotPassModel
                    turnCount = 0
                    
                    'tant qu'il n'y a pas de vainqueur et qu'on a joué moins de 300 tours
                    While Not board.IsWinner And turnCount < 300 And deamonReapled < 3
                    
                        Call ysnp.Snapshot
                        
                        If board.CurrentTurn = EColor.White Then
                            Call Bot.Run
                        ElseIf board.CurrentTurn = EColor.Black Then
                            Call Bot.Run
                        End If
                        
                        If ysnp.IsSuccess Then
                            deamonReapled = 0
                            turnCount = turnCount + 1
                            Call board.SwitchTurn
                            Call board.SearchWinner
                        Else
                            deamonReapled = deamonReapled + 1
                            Call ysnp.RollBack()
                        End If
                        
                    Wend
                    
                    If deamonReapled = 3 Then
                        MsgBox "You Failed"
                    End If
                    
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
            Call BoardConstructeur.FormatBoard
            'initialise les pions à leur position de départ
            Call BoardConstructeur.Initialisation
            
    End Select
End Sub

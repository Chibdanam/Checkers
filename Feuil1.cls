VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'/// PROC�DURE  : Se d�clenche lorsqu'un double clic survient sur la feuille et transmet en parametre la range de la cellule cliqu�e
'/// PARAM�TRE  : Range, Boolean
'/// RETOUR     : Aucun (evenement)
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

  'Si Cancel a sa valeur par défaut ou égal à False, alors après la procédure, 
  'Excel va continuer en traitant le clic droit comme il le fait par défaut. 

  'Si Cancel est à True, alors Excel ne va pas faire son traitement habituel, 
  'Excel considère que l'événement du clic droit est entièrement pris en charge par la procédure.
  Cancel = true

    Dim checkerBoard As BoardModel
    Dim pawn As PawnModel
    Dim enemyPawn As PawnModel
    
    'on instancie un object correspondant a notre plateau de jeu
    Set checkerBoard = New BoardModel
   
    'selon la section dans laquelle on a cliqu�
    Select Case checkerBoard.GetSection(Target)
    
        'si on a cliqu� sur le damier
        Case ESection.Game
        
            'selon la valeur actuelle de configuration de la partie
            Select Case checkerBoard.ConfigPlayer
            
                'si en mode 2 joueurs
                Case EConfig.TwoPlayer
                    
                    'si l'action correspond au mouvement d'un pion
                    If Player.Run(Target) Then
                    
                        'v�rifie s'il y a un vainqueur
                        Call checkerBoard.SearchWinner
                        
                        'change la couleur afficher sur le plateau
                        Call checkerBoard.SwitchTurn
                    
                    End If
                    
                'si en mode 1 joueur
                Case EConfig.SinglePlayer
                
                    'si l'action correspond au mouvement d'un pion
                    If Player.Run(Target) Then
                    
                        'v�rifie s'il y a un vainqueur
                        Call checkerBoard.SearchWinner
                        
                        'change la couleur afficher sur le plateau
                        Call checkerBoard.SwitchTurn
                        
                        's'il n'y a pas de vainqueur
                        If Not checkerBoard.IsWinner Then
                        
                            'fais jouer un tour au bot
                            Call Bot.Run
                            
                            'v�rifie s'il y a un vainqueur
                            Call checkerBoard.SearchWinner
                            
                            'change la couleur afficher sur le plateau
                            Call checkerBoard.SwitchTurn
                            
                        End If
                        
                    End If
                
                'si en mode IA vs IA
                Case EConfig.Automate
                    
                    Dim turnCount As Integer
                    turnCount = 0
                    
                    'tant qu'il n'y a pas de vainqueur et qu'on a jouer moins de 300 tour
                    While Not checkerBoard.IsWinner And turnCount < 300
                    
                        'si la couleur du joueur en cours est blanc
                        If checkerBoard.CurrentTurn = EColor.White Then
                        
                            'fais jouer un bot
                            Call Bot.Run
                            'change la couleur afficher sur le plateau
                            Call checkerBoard.SwitchTurn
                            
                        'si la couleur du joueur en cours est blanc
                        ElseIf checkerBoard.CurrentTurn = EColor.Black Then
                        
                            'fais jouer un bot
                            Call Bot.Run
                            'change la couleur afficher sur le plateau
                            Call checkerBoard.SwitchTurn
                            
                        End If
                        
                        'v�rifie s'il y a un vainqueur
                        Call checkerBoard.SearchWinner
                        'on incr�mente le nombre de tour
                        turnCount = turnCount + 1
                    Wend
                    
            End Select
            
                
        'cas de clic sur la configuration des joueurs, on incr�mente la configuration
        Case ESection.ConfigPlayer
        
            'si la configuration actuelle est 1 joueurs
            If checkerBoard.ConfigPlayer = EConfig.SinglePlayer Then
                'on passe en configuration 2 joueurs
                checkerBoard.ConfigPlayer = EConfig.TwoPlayer
                
            'si la configuration actuelle est 2 joueurs
            ElseIf checkerBoard.ConfigPlayer = EConfig.TwoPlayer Then
                'on passe en configuration IA vs IA
                checkerBoard.ConfigPlayer = EConfig.Automate
                
            'si la configuration actuelle n'est ni 1 joueurs, ni 2 joueurs
            Else
                checkerBoard.ConfigPlayer = EConfig.SinglePlayer
            End If
        
        'cas de clic sur le bouton restart
        Case ESection.Restart
            'formate le plateau de jeux sur la feuille excel
            Call BoardConstructeur.FormatBoard
            'initialise les pions a leur position de d�part
            Call BoardConstructeur.Initalisation
            
    End Select
End Sub


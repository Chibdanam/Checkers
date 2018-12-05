Attribute VB_Name = "BotManager"
Option Explicit
Option Base 0



'/// PROCÉDURE  : Récupère le bot en fonction de la couleur du tour puis l'exectue
'/// PARAMÈTRE  : EColor
'/// RETOUR     : Aucun
Public Sub Run(pColor As EColor)
Dim bot As String

    If EnumString(pColor) = "White" Then
        bot = "Bot.Run"
    ElseIf EnumString(pColor) = "Black" Then
        bot = "Bot.Run2"
    End If

    Call RunBot(bot)

End Sub



'/// PROCÉDURE  : Execute le bot et vérifie qu'il n'a enfreint aucunes règles
'/// PARAMÈTRE  : String
'/// RETOUR     : Aucun
Private Sub RunBot(pBotName As String)
Dim ysnp As YouShallNotPassModel
Dim deamonReapled As Integer
Dim validLap As Boolean

    'instancie le vérificateur
    Set ysnp = New YouShallNotPassModel

    'sauvegarde la configuration du plateau avant le tour du bot
    Call ysnp.Snapshot

    'tant que le bot n'a pas effectué un tour correct, et tant qu'il n'a pas effectuer 3 echec
    While Not validLap And deamonReapled < 3
        
        Application.Run pBotName

        If ysnp.IsSuccess Then
            'on valide le tour
            validLap = True
        Else
            'on incrémente le compteur d'erreur
            deamonReapled = deamonReapled + 1
            'on restaure le plateau avant l'action du bot
            Call ysnp.Rollback
        End If
    Wend
    
    If deamonReapled = 3 Then
        Range("TurnValue") = pBotName + " failed"
        MsgBox pBotName + " failed"
    End If
    
End Sub

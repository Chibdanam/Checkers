Attribute VB_Name = "BotManager"
Option Explicit
Option Base 0
Private prvWhiteBot As String
Private prvBlackBot As String



'/// PROC�DURE  : R�cup�re le bot blanc
'/// PARAM�TRE  : Aucun
'/// RETOUR     : String
Public Function GetWhiteBot() As String
    GetWhiteBot = prvWhiteBot
End Function



'/// PROC�DURE  : R�cup�re le bot noir
'/// PARAM�TRE  : Aucun
'/// RETOUR     : String
Public Function GetBlackBot() As String
    GetBlackBot = prvBlackBot
End Function



'/// PROC�DURE  : Initialise les bots par defaut
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Public Sub InitBot()
    If prvWhiteBot = "" Then prvWhiteBot = "FirstBot"
    If prvBlackBot = "" Then prvBlackBot = "Random"
    
    Sheets("BOARD").Shapes("lblDisplayWhiteBot").TextFrame.Characters.Text = "W - " + prvWhiteBot
    Sheets("BOARD").Shapes("lblDisplayBlackBot").TextFrame.Characters.Text = "B - " + prvBlackBot
End Sub



'/// PROC�DURE  : R�cup�re le bot en fonction de la couleur du tour puis l'ex�cute
'/// PARAM�TRE  : EColor
'/// RETOUR     : Aucun
Public Sub Run(pColor As EColor)
Dim bot As String

    If EnumString(pColor) = "White" Then
        bot = GetWhiteBot
    ElseIf EnumString(pColor) = "Black" Then
        bot = GetBlackBot
    End If

    Call RunBot(bot)

End Sub


'/// PROC�DURE  : Ex�cute le bot et v�rifie qu'il n'a enfreint aucunes r�gles
'/// PARAM�TRE  : String
'/// RETOUR     : Aucun
Private Sub RunBot(pBotName As String)
Dim ysnp As YouShallNotPassModel
Dim deamonReapled As Integer
Dim confirmLap As Boolean
Dim lapTimer As Single

    'instancie le v�rificateur
    Set ysnp = New YouShallNotPassModel

    'sauvegarde la configuration du plateau avant le tour du bot
    Call ysnp.Snapshot

    lapTimer = 0

    'tant que le bot n'a pas effectu� un tour correct
    While Not confirmLap
        
        lapTimer = Timer

        Application.Run "Bot_" + pBotName + ".Run"
        
        'on calcule le temps d'ex�cution du bot
        lapTimer = (Timer - lapTimer) * 1000

        Debug.Print (pBotName + ": " + CStr(lapTimer) + "ms")

        If ysnp.IsSuccess Then
            'on valide le tour
            confirmLap = True
            Call Log.CG_UpdateTurnDuration(lapTimer)
        Else
            'on incr�mente le compteur d'erreur
            deamonReapled = deamonReapled + 1
            
            'on restaure le plateau avant l'action du bot
            Call ysnp.Rollback
        End If
        
        If deamonReapled = 3 Or lapTimer > 5000 Then
            confirmLap = True
            Range("TurnValue") = pBotName + " failed"
            MsgBox "Bot Failed" + vbNewLine + _
                   "Name       : " + pBotName + vbNewLine + _
                   "Wrong move : " + CStr(deamonReapled) + vbNewLine + _
                   "Time lap   : " + CStr(lapTimer)

        End If
    Wend
    
End Sub

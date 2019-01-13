Attribute VB_Name = "BotManager"
Option Explicit
Option Base 0



'/// PROC�DURE  : Initialise les bots par defaut
'/// PARAM�TRE  : Aucun
'/// RETOUR     : Aucun
Public Sub InitBot()
Dim board As BoardModel
    Set board = New BoardModel
    If board.WhiteBot = "" Then board.WhiteBot = "Default"
    If board.BlackBot = "" Then board.BlackBot = "Default"
End Sub



'/// PROC�DURE  : R�cup�re le bot en fonction de la couleur du tour puis l'ex�cute
'/// PARAM�TRE  : EColor
'/// RETOUR     : Aucun
Public Sub Run(pColor As EColor)
Dim bot As String
Dim board As BoardModel

    Set board = New BoardModel

    If EnumString(pColor) = "White" Then
        bot = board.WhiteBot
    ElseIf EnumString(pColor) = "Black" Then
        bot = board.BlackBot
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
        
        On Error Resume Next
        Application.Run "Bot_" + pBotName + ".Run"
        On Error GoTo 0
        
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

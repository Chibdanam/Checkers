VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFOptions 
   Caption         =   "Options"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UFOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton_Replay_Click()
    Application.ScreenUpdating = False
    Call StartReplay_CurrentGame
    Application.ScreenUpdating = True
End Sub

    
Private Sub StartReplay_CurrentGame()
    
    If OptionButton_LastGame.Value = True Then
        Set turns = Worksheets("CURRENT GAME").ListObjects("CURRENT_TURNS_DATA").ListRows
        Call ReplayGame(turns)
    ElseIf OptionButton_SelectGame.Value = True Then
    
    End If

End Sub


Private Sub ReplayGame(ByVal pTurns As Variant)
Dim turnID As Integer

    If Not Tools.IsArrayNullOrEmpty(pTurns) Then
        
        UFOptions.Hide
        
        For Each turn In pTurns
            
            turnID = turn.Range.Cells(1, 1).Value
            If turnID = 1 Then
                
                blueprint = GetRangeFromTable(pTurns.Parent.Parent.Name, pTurns.Parent.Name, "Board initial state", turnID).Value
                Call Compute(blueprint)
                
            End If
            
            blueprint = GetRangeFromTable(pTurns.Parent.Parent.Name, pTurns.Parent.Name, "Board final state", turnID).Value
            Call Compute(blueprint)
            
            Call Tools.RefreshScreen(CInt(TextBox_Sleep.Value))
            
        Next turn
        
    End If

End Sub

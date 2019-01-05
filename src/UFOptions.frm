VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFOptions 
   Caption         =   "Options"
   ClientHeight    =   3036
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4590
   OleObjectBlob   =   "UFOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub UserForm_Initialize()
'    Dim gamesCount As Integer
'
'    For i = 1 To 5
'        ComboBox1.AddItem "Ligne" & i
'    Next i
    ' Affecter une valeur par défaut lors de l'affichage du ComboBox.
'    ComboBox1.ListIndex = 0
End Sub


Private Sub CommandButton_Replay_Click()
    Application.ScreenUpdating = False
    Call StartReplay
    Application.ScreenUpdating = True
End Sub

    
Private Sub StartReplay()
    
    If OptionButton_LastGame.Value = True Then
        Set Turns = Worksheets("CURRENT GAME").ListObjects("CURRENT_TURNS_DATA").ListRows
        Call ReplayGame(Turns)
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

        Unload UFOptions
        
    End If

End Sub

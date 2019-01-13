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
Private Sub CommandButton_ApplyConfig_Click()
Dim module As Variant
Dim whiteGood As Boolean
Dim blackGood As Boolean

    For Each module In ThisWorkbook.VBProject.VBComponents
        If module.Type = 1 And Left(module.Name, 4) = "Bot_" Then
            If ComboBox_WhiteBot.Value = Right(module.Name, Len(module.Name) - 4) Then whiteGood = True
            If ComboBox_WhiteBot.Value = Right(module.Name, Len(module.Name) - 4) Then blackGood = True
        End If
    Next module
    
    If whiteGood And blackGood Then
        Sheets("BOARD").Shapes("lblDisplayWhiteBot").TextFrame.Characters.Text = "W - " + ComboBox_WhiteBot.Value
        Sheets("BOARD").Shapes("lblDisplayBlackBot").TextFrame.Characters.Text = "B - " + ComboBox_BlackBot.Value
        Unload Me
    Else
        MsgBox "Please select a bot from the list"
    End If
    
End Sub

Private Sub UserForm_Initialize()
Dim gamesCount As Integer
Dim module As Variant

    For Each module In ThisWorkbook.VBProject.VBComponents

        'if normal module
        If module.Type = 1 And Left(module.Name, 4) = "Bot_" Then
            ComboBox_WhiteBot.AddItem Right(module.Name, Len(module.Name) - 4)
            ComboBox_BlackBot.AddItem Right(module.Name, Len(module.Name) - 4)
        End If
        
    Next module
    
    whiteLabel = Sheets("BOARD").Shapes("lblDisplayWhiteBot").TextFrame.Characters.Text
    blackLabel = Sheets("BOARD").Shapes("lblDisplayBlackBot").TextFrame.Characters.Text
    
    'Affecter une valeur par défaut lors de l'affichage du ComboBox.
    If IsEmpty(Right(whiteLabel, Len(whiteLabel) - 4)) Then
        ComboBox_BlackBot.ListIndex = 0
    Else
        ComboBox_WhiteBot.Value = Right(whiteLabel, Len(whiteLabel) - 4)
    End If
    
    If IsEmpty(Right(blackLabel, Len(blackLabel) - 4)) Then
        ComboBox_BlackBot.ListIndex = 0
    Else
        ComboBox_BlackBot.Value = Right(blackLabel, Len(blackLabel) - 4)
    End If
    
End Sub


Private Sub CommandButton_StartReplay_Click()
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

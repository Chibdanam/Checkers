VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "YouShallNotPassModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private prvInitialWhitePawns As Variant     '/// VARIABLE PRIVÉE : Liste de pion allié initiale
Private prvInitialBlackPawns As Variant     '/// VARIABLE PRIVÉE : Liste de pion enemi initiale
Private prvSavedSnapshot As String               '/// VARIABLE PRIVÉE : String représentant une configuration du plateau



'/// CONSTRUCTEUR: procédure appelé lors de l'instanciation de l'objet
'/// PARAMÈTRE   : Aucun
'/// RETOUR      : Aucun
Private Sub class_Initialize()
End Sub



'/// DESTRUCTEUR: procédure appelé lors de la destruction de l'objet
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Private Sub class_Terminate()
End Sub



'/// PROCÉDURE  : Enregistre l'etat actuel du jeu
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub Snapshot()
    prvInitialWhitePawns = Tools.GetPawns(EColor.White)
    prvInitialBlackPawns = Tools.GetPawns(EColor.Black)
    prvSavedSnapshot = Tools.MakeBlueprintFromBoard
End Sub



'/// PROCÉDURE  : Restore la snapshot en tant que configuration de la partie
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Aucun
Public Sub Rollback()
    Call Tools.Compute(prvSavedSnapshot)
End Sub



'/// PROPRIÉTÉ  : Vérifie si la différence entre le snapshot et l'etat du jeu actuel correspond à un coup suivant les règle
'/// PARAMÈTRE  : Aucun
'/// RETOUR     : Boolean
Public Property Get IsSuccess() As Boolean
Dim board As BoardModel
Dim alliesInitial As Variant
Dim alliesFinal As Variant
Dim enemiesInital As Variant
Dim enemiesFinal As Variant
Dim missingAllies As Variant
Dim newAllies As Variant
Dim currentSnapshot As String
Dim mockSnapshot As String
Dim originalPawn As PawnModel
Dim originalTarget As PawnModel

    Application.ScreenUpdating = False

    Set board = New BoardModel
    
    If board.TurnColor = EColor.White Then
        alliesInitial = prvInitialWhitePawns
        alliesFinal = Tools.GetPawns(EColor.White)
    Else
        alliesInitial = prvInitialBlackPawns
        alliesFinal = Tools.GetPawns(EColor.Black)
    End If
    
    missingAllies = GetMissmatchedPawnsFromSnapshot(alliesInitial, alliesFinal, True)
    newAllies = GetMissmatchedPawnsFromSnapshot(alliesInitial, alliesFinal, False)
    
    currentSnapshot = Tools.MakeBlueprintFromBoard
    
    If UBound(missingAllies) = 0 And UBound(newAllies) = 0 Then
        Call Tools.Compute(prvSavedSnapshot)
        
        Set originalPawn = New PawnModel
        Call originalPawn.Build(missingAllies(0).CurrentRange)
        
        Set originalTarget = New PawnModel
        Call originalTarget.Build(newAllies(0).CurrentRange)
        
        If originalPawn.TryMoveTo(originalTarget, True) Then
            mockSnapshot = Tools.MakeBlueprintFromBoard
            If mockSnapshot = currentSnapshot Then
                IsSuccess = True
            End If
        End If
    End If
    
    Application.ScreenUpdating = True
    
End Property



'/// FONCTION   : Cherche les pions manquant ou en trop dans une liste de pion par rapport a une autre liste de pion
'/// PARAMÈTRE  : Variant, Variant, Boolean
'/// RETOUR     : Variant
Private Function GetMissmatchedPawnsFromSnapshot(ByVal pPawnsInitial As Variant, ByVal pPawnsFinal As Variant, ByVal pMissing As Boolean) As Variant
Dim missmatchedPawns() As PawnModel
Dim pawnsToFind As Variant
Dim pawn As Variant
Dim pawnsCandidate As Variant
Dim pawnCandidate As Variant
Dim missmatchedPawnsCount As Integer
Dim match As Boolean
        
    missmatchedPawnsCount = 0
    
    If pMissing Then
        pawnsToFind = pPawnsInitial
        pawnsCandidate = pPawnsFinal
    Else
        pawnsToFind = pPawnsFinal
        pawnsCandidate = pPawnsInitial
    End If
    
    For Each pawn In pawnsToFind
        match = False
        For Each pawnCandidate In pawnsCandidate
            If pawn.CurrentRange.Address = pawnCandidate.CurrentRange.Address And _
            pawn.CurrentRange.Value = pawnCandidate.CurrentRange.Value Then
                match = True
                Exit For
            End If
        Next pawnCandidate
        If Not match Then
            ReDim Preserve missmatchedPawns(missmatchedPawnsCount)
            Set missmatchedPawns(missmatchedPawnsCount) = pawn
            missmatchedPawnsCount = missmatchedPawnsCount + 1
        End If
    Next pawn
    
    GetMissmatchedPawnsFromSnapshot = missmatchedPawns

End Function

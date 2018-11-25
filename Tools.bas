Attribute VB_Name = "Tools"
'/// https://www.exceltrick.com/formulas_macros/vba-wait-and-sleep-functions/?
'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'/// FONCTION   : Retourne un tableau contenant tous les pions de la couleur pass�e en param�tre
'/// PARAM�TRE  : EColor
'/// RETOUR     : Variant (tableau de pion)
Public Function GetPawns(pColor As EColor) As Variant

    Dim pawnList() As PawnModel
    Dim pawnCounter As Integer
    pawnCounter = 0
    
    'pour chaque cellule du damier
    For Each cell In Range("Game")
        
        'on instancie un objet pion et on le construit avec la range de la cellule actuelle
        Set pawnCandidate = New PawnModel
        Call pawnCandidate.Build(cell.Cells(1, 1))
        
        'si le pion est v�ritablement un pion et est de la couleur pass� en param�tre
        If pawnCandidate.IsPawn And pawnCandidate.Color = pColor Then
            
            'on redimensionne notre tableau
            ReDim Preserve pawnList(pawnCounter)
            'on ajoute le pion au tableau de pion
            Set pawnList(pawnCounter) = pawnCandidate
            'on incr�mente notre compteur
            pawnCounter = pawnCounter + 1
            
        End If
        
    Next cell
    
    'on associe le tableau de pion ainsi constitu� au retour de la fonction
    GetPawns = pawnList
    
End Function



'/// FONCTION   :
'/// PARAM�TRE  :
'/// RETOUR     :
Function IsArrayNullOrEmpty(arr As Variant) As Boolean
    
    On Error Resume Next
    
    IsArrayNullOrEmpty = True
    
    If IsArray(arr) And _
    Not IsError(LBound(arr, 1)) And _
    LBound(arr, 1) <= UBound(arr, 1) And _
    Not IsEmpty(arr) And _
    UBound(arr) > -1 Then
        IsArrayNullOrEmpty = False
    End If
    
    Exit Function
    
End Function



'/// FONCTION   :
'/// PARAM�TRE  :
'/// RETOUR     :
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function


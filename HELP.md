# Aide

<!-- TOC -->

- [Instructions de module](#instructions-de-module)
  - [Option Base](#Option-Base)
  - [Option Explicit](#Option-Explicit)
- [Gestion des Arrays](#Gestion-des-Arrays)
  - [UBound](#UBound)
  - [LBound](#LBound)
  - [Redim](#Redim)
- [Gestion d'erreurs](#Gestion-d'erreurs)
  - [IsError](#IsError)
- [Enumération](#Enumération)
    - [Enum](#Enum)
- [Manipulation de chaines de caractères](#Manipulation-de-chaines-de-caractères)
  - [Left](#Left)
  - [Right](#Right)
  - [Mid](#Mid)
  - [Split](#Split)
  - [Trim](#Trim)
- [Module de Classe](#Module-de-Classe)
  - [class_Initialize et class_Terminate](#class_Initialize-et-class_Terminate)
- [Methodes](#Methodes)
  - [Public](#Public)
  - [Private](#Private)
  - [ByVal](#ByVal)
  - [ByRef](#ByRef)
  - [Property Get / Let](#Property-Get-/-Let)
- [Instructions](#Instructions)
  - [Call](#Call)
  - [Set](#Set)
  - [New](#New)
  - [With](#With)
  - [Select Case](#Select-Case)
  - [For Each](#For-Each)
- [Application](#Application)
  - [ActiveWorkbook.Names.Add](#ActiveWorkbook.Names.Add)
  - [Debug](#Debug)

<!-- /TOC -->

## Instructions de module

### Option Base
Déclare la limite inférieure par défaut des indices de tableau
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/option-base-statement)  
``` 
Option Base 1 ' Set default array subscripts to 1                    
 
Dim Lower 
Dim MyArray(20)
Dim ZeroArray(0 To 5) ' Override default base subscript          

' Use LBound function to test lower bounds of arrays
Lower = LBound(MyArray)     ' Returns 1
Lower = LBound(ZeroArray)   ' Returns 0 
```

### Option Explicit
Impose la déclaration explicite de toutes les variables du module
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/option-explicit-statement)
``` 
Option Explicit ' Force explicit variable declaration

Dim counter

For i = 1 to 10 ' Undeclared variable i generate error
For counter = 1 to 10  ' Declared variable does not generate error
```
## Gestion des Arrays

### UBound
Renvoie le plus grand indice disponible pour une dimension d'un tableau.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/ubound-function)
``` 
Dim Upper As Integer
Dim MyArray(10)

Upper = UBound(MyArray)    ' Returns 10
```

### LBound
Renvoie le plus petit indice disponible pour une dimension d'un tableau.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/lbound-function)
``` 
Dim Lower
Dim MyArray(10)

Lower = Lbound(MyArray)     ' Returns 0 or 1, depending on setting of Option Base
```

### Redim 
Redimensionne les variables de [tableau dynamique](https://silkyroad.developpez.com/vba/tableaux/#LII-B). Afin de conserver les variables déjà présentes dans le tableau, utiliser le mot clef *Preserve*.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/redim-statement)
``` 
Option Base 1

Dim MyArray() As Integer ' Declare dynamic array of Integer

Redim MyArray(5) ' Allocate 5 elements

For i = 1 To 5      ' Loop 5 times
 MyArray(i) = i     ' Insert value 
Next i

Redim MyArray(10) ' Create new array over the last and resize to allocate 10 elements

For i = 1 To 10     ' Loop 10 times
 MyArray(i) = i     ' Initialize array
Next i

Redim Preserve MyArray(15) ' Resize to 15 elements keeping last 10 values
```

## Gestion d'erreurs

### IsError
Renvoie un booléen indiquant si une expression retourne une erreur.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/iserror-function)
``` 
IsError(2/0) 'Return true
```

## Enumération

### Enum
Les variables d’énumération sont des variables déclarées avec un type Enum. Les éléments du type Enum sont initialisés avec des valeurs constantes
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/enum-statement)
``` 
Enum EWindRose
    NorthEast
    NorthWest
    SouthEast
    SouthWest
End Enum

Dim myDirection As EWindRose

myDirection = NorthEast
```

## Manipulation de chaines de caractères

### Left
Renvoie un Variant (String) contenant un nombre spécifique de caractères dans la partie gauche d'une chaîne.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/left-function)
``` 
Dim myStr

myStr = Left("Hello World", 1)   ' Returns "H"
myStr = Left("Hello World", 7)   ' Returns "Hello W"
myStr = Left("Hello World", 20)  ' Returns "Hello World"
```

### Right
Renvoie une valeur de type Variant (String) contenant un nombre spécifié de caractères à partir de l’extrémité droite d’une chaîne.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/right-function)
``` 
Dim myStr

myStr = Right("Hello World", 1)    ' Returns "d"
myStr = Right("Hello World", 5)    ' Returns "World"
myStr = Right("Hello World", 20)   ' Returns "Hello World"
```

### Mid
Retourne une valeur de type Variant ( String ) contenant un nombre indiqué de caractères extraits d'une chaîne de caractères.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/mid-function)
``` 
Dim myStr

myStr = Mid("Hello World", 1, 5)   ' Returns "Hello".
myStr = Mid("Hello World", 6, 5)   ' Returns "World".
myStr = Mid("Hello World", 5, 3)   ' Returns "o W".
```


### Split
Renvoie un tableau unidimensionnel de base zéro contenant un nombre spécifié de sous-chaînes.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/split-function)
``` 
Split("Ecam Strasbourg - Europe")
Result: {"Ecam", "Strasbourg","-", "Europe" }

Split("03.88.45.45", ".")
Result: {"03", "88", "45", "45"}

Split("A;B;C;D", ";")
Result: {"A", "B", "C", "D"}
```
```
Dim myStr As String
Dim myStrs() As String

myStr = "Ecam Strasbourg - Europe"
myStrs = Split(myStr,"-")

Debug.Print myStrs(0) 'return "Ecam Strasbourg "
Debug.Print myStrs(1) 'return " Europe"
```


### Trim
Renvoie une valeur de type Variant (String) contenant une copie d’une chaîne en supprimant les espaces de gauche (LTrim), les espaces de droite (RTrim) ou les deux (Trim).
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/ltrim-rtrim-and-trim-functions)
``` 
Dim TrimString

TrimString = LTrim("  <-Trim->  ")    ' TrimString = "<-Trim->  "
TrimString = RTrim("  <-Trim->  ")    ' TrimString = "  <-Trim->"
TrimString = Trim("  <-Trim->  ")     ' TrimString = "<-Trim->"
```

## Module de Classe 

### class_Initialize et class_Terminate

[documentation sur les modules de classe](https://sinarf.developpez.com/access/vbaclass/#L2.4)


## Methodes

### Public
Public - Indique que la procédure est accessible à toutes les autres procédures dans tous les modules. Si elles ne sont pas explicitement spécifiées avec Public ou Private les procédures sont publiques par défaut
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/public-statement)
``` 
Public Number As Integer    ' Public Integer variable.
```

### Private
Private - Indique que la procédure est uniquement accessible aux autres procédures du module dans lequel elle est déclarée
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/private-statement)
``` 
Private Number As Integer    ' Private Integer variable.
```

### ByVal
ByVal Indique que l'argument est transmis par valeur.
[doc](https://docs.microsoft.com/en-gb/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference)  
``` 
Sub TestByVal()
Dim number As Integer
Dim result As Integer

    number = 3                      
    result = Computer(number)

    Debug.Print(Cstr(number))       'print 3
    Debug.Print(Cstr(result))       'print 7

End Sub

Function Computer(ByVal pNumber As Integer)

    pNumber = pNumber + 4           'modify pNumber value in Computer function
    Computer = pNumber

End Function
```

### ByRef
ByRef Indique qu'un argument est transmis par référence. L'élément ByRef est la valeur par défaut dans Visual Basic.
[doc](https://docs.microsoft.com/en-gb/dotnet/visual-basic/programming-guide/language-features/procedures/passing-arguments-by-value-and-by-reference)  
``` 
Sub TestByRef()
Dim number As Integer
Dim result As Integer

    number = 3                      
    result = Compute(number)

    Debug.Print(Cstr(number))       'print 7
    Debug.Print(Cstr(result))       'print 7

End Sub

Function Computer(ByRef pNumber As Integer)
    pNumber = pNumber + 4           'modify pNumber value in Computer function
    Computer = pNumber
End Function
```


### Property Get / Let

Property **Get** permet la lecture d'une propriété d'un module de classe
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/property-get-statement)
``` 
Private prvNom As String

Property Get Nom() As String
    ' Propriété en lecture
    Nom = prvNom
End Property
```


Property **Let** permet l'écriture d'une propriété d'un module de classe
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/property-let-statement)
``` 
Private prvNom As String

Property Let Nom(pNom As String)
    ' Propriété en écriture
    prvNom = pNom
End Property
```

## Instructions

### Call
Transfère le contrôle à une procédure Sub , une procédure Function
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/call-statement)
``` 
Call MySub
```

### Set 
Attribue une référence d’objet à une variable ou à une propriété
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/set-statement)
``` 
Dim myRange As Range
Set myRange = Range("A1")
```
### New 
Permet de créer une nouvelle instance d'un objet.
[doc](https://stackoverflow.com/questions/21652671/what-does-the-keyword-new-do-in-vba)
``` 
Dim myObj As Object
Set myObj = New Object
```

### With
L’instruction With vous permet d’effectuer une série d’instructions sur un objet spécifié sans qualifier le nom de l’objet à chaque fois.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/with-statement)
``` 
With Range("A1")
    .RowHeight = 25
    .ColumnWidth = 4
    .Font.Size = 14
    .Font.Bold = True
End With
```

### Select Case
Exécute un ou plusieurs groupes d' instructions, selon la valeur d'une expression.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/select-case-statement)
``` 
Dim Number 
Number = 8    

Select Case Number     
    Case 1    
        Debug.Print "1" 
    Case 2    
        Debug.Print "2" 
    Case 3    
        Debug.Print "3" 
    Case Else    ' Other values. 
        Debug.Print "Not between 1 and 3" 
End Select
```

### For Each
Répète un groupe d'instructions pour chaque élément dans un Variant ou une Collection
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/for-eachnext-statement)
``` 
Dim day as Variant
Dim week as Variant

week = Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")

For Each day In week
    Debug.Print(day)
Next day
```
```
Dim myCell As Range

For Each myCell In Range("A1:Z26")
    Debug.Print(myCell.Address)
Next day
```

## Application

### ActiveWorkbook.Names.Add
Ajoute une plage nommée au classeur
[doc](https://docs.microsoft.com/en-gb/office/vba/api/excel.names.add)
``` 
ActiveWorkbook.Names.Add Name:="Game", RefersToR1C1:="=Feuil1!R2C2:R9C9"

Range("Game")   'refers to Range("B2:I9")
```

### Debug
L'objet Debug permet d'afficher des valeurs dans la fentre immediate lors du déboggage
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/debug-object)
``` 
Debug.Print "Hello Immediate window"
```

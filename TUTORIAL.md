## Summary

- [Class Module](#Class-Module)
- [Module](#Module)
- [Class Module : PawnModel](#Class-Module-:-PawnModel)
- [Class Module : MoveModel](#Class-Module-:-MoveModel)
- [Class Module : BoardModel](#Class-Module-:-BoardModel)
- [Module : Tools](#Module-:-Tools)
    - [MakeBlueprintFromBoard](#MakeBlueprintFromBoard)
    - [Compute](#Compute)
    - [GetPawns](#GetPawns)

## Class Module
To use the functions of a class module in VBA it is necessary to
- Declare the variable representing the object ```Dim maVariable As Object```
- Associate an instance to our variable with the instruction ```Set maVariable = uneInstance```, or create a new instance with the instruction ```Set maVariable = New Object```
- Write ```myVariable.Mafunction```

```
Dim maVariable
Dim monInstance As ClassModule1
Set monIntance = New ClassModule1
maVariable = monIntance.MaFonction()
```

## Module

To call functions of a module, just call the function directly by its name
```
Dim maVariable
maVariable = maFonction()
```
However, a function can be written in two different modules and with the same name, it is therefore necessary to specify systematically the module where is written the function called
``` 
Dim maVariable
maVariable = Module1.maFonction()
```

## Class Module : PawnModel

The PawnModel object represents a pawn of the board. It is built from a Range.

instantiation and construct on Range("C2")
``` 
Dim pawn As PawnModel

Set pawn = New PawnModel
Call pawn.Build(Range("C2"))
```

Once the object is instantiate, you can read or write its property
```
Dim canMove As Boolean

canMove = pawn.CanMove 'Return true or false
```
Or call its function, here an exemple for moving the pawn to cell Range("D3")
```
Dim hasMoved As Boolean
Dim target As PawnModel

Set target = New PawnModel
Call target.Build(Range("D3"))

hasMoved = pawn.TryMoveTo(target, true) 'return true if the pawn has moved to target
```

## Class Module : MoveModel

The MoveModel object represents the movement of a pawn to a cell. It is built from a pawn and a Range.

instantiation and construct with Pawn on Range("C2") and target on Range("D3")
``` 
Dim pawn As PawnModel
Dim target As PawnModel
Dim move As MoveModel

Set pawn = New PawnModel
Call pawn.Build(Range("C2"))

Set target = New PawnModel
Call target.Build(Range("D3"))

Set move = New MoveModel
Call pawn.Build(pawn,target)
```

## Class Module : BoardModel
The BoardModel object represents the board. Since there is only one board there is no need to build it, all instances of BoardModel return the same board object. If we modify its parameters, then all instances of BoardModel will have their parameters modified simultaneously

instantiation 
``` 
Dim board As BoardModel
Set board = New BoardModel
```

exemple : ```board.TurnColor``` récupère la couleur du tour


## Module : Tools

### MakeBlueprintFromBoard
This procedure saves the board as a string  
w = white pawn  
W = white queen  
b = black pawn  
B = black queen
``` 
Dim snapshot As String
Dim board As BoardModel
Set board = New BoardModel

Call board.Initialisation

snapshot = Tools.MakeBlueprintFromBoard

'snapshot ="|-|b|-|b|-|b|-|b|
'           |b|-|b|-|b|-|b|-|
'           |-|b|-|b|-|b|-|b|
'           | |-| |-| |-| |-|
'           |-| |-| |-| |-| |
'           |w|-|w|-|w|-|w|-|
'           |-|w|-|w|-|w|-|w|
'           |w|-|w|-|w|-|w|-|"
```

### Compute
Compute allows you to print a string on the board
``` 
Dim blueprint As String

blueprint = "|-|b|-|b|-|b|-|b|" + vbNewLine + _
            "|b|-|b|-|b|-|b|-|" + vbNewLine + _
            "|-|b|-|b|-|b|-|b|" + vbNewLine + _
            "| |-| |-| |-| |-|" + vbNewLine + _
            "|-| |-| |-| |-| |" + vbNewLine + _
            "|w|-|w|-|w|-|w|-|" + vbNewLine + _
            "|-|w|-|w|-|w|-|w|" + vbNewLine + _
            "|w|-|w|-|w|-|w|-|"
            
Call Tools.Compute(blueprint)
```

### GetPawns
GetPawns allows to recover all the pieces of a color on the board
``` 
Dim pawnList As Variant
Dim pawn As PawnModel
Dim pawnListCount As Integer
Dim rangeOfSelectedPawn As Range

pawnList = Tools.GetPawns(EColor.White)
pawnListCount = UBound(pawnList) 'return the number of pawn in the list

Set pawn = pawnList(2) 'associate the second pawn in the list with the pawn variable
Set rangeOfSelectedPawn = pawn.CurrentRange 'return the range of the selected pawn
```
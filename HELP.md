# Summary

<!-- TOC -->

- [Statements](#Statements)
  - [Option Base](#Option-Base)
  - [Option Explicit](#Option-Explicit)
- [Arrays Management](#Arrays-Management)
  - [UBound](#UBound)
  - [LBound](#LBound)
  - [Redim](#Redim)
- [Error management](#Error-management)
  - [IsError](#IsError)
- [Enumeration](#Enumeration)
    - [Enum](#Enum)
- [Manipulation of strings](#Manipulation-of-strings)
  - [Left](#Left)
  - [Right](#Right)
  - [Mid](#Mid)
  - [Split](#Split)
  - [Trim](#Trim)
- [Class Module](#Class-Module)
  - [class_Initialize and class_Terminate](#class_Initialize-and-class_Terminate)
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

## Statements

### Option Base
Declares the default lower limit for array indices
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
Impose the explicit declaration of all the variables of the module
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/option-explicit-statement)
``` 
Option Explicit ' Force explicit variable declaration

Dim counter

For i = 1 to 10 ' Undeclared variable i generate error
For counter = 1 to 10  ' Declared variable does not generate error
```
## Arrays Management

### UBound
Returns the largest index available for a dimension of a table.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/ubound-function)
``` 
Dim Upper As Integer
Dim MyArray(10)

Upper = UBound(MyArray)    ' Returns 10
```

### LBound
Returns the smallest index available for a dimension of a table.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/lbound-function)
``` 
Dim Lower
Dim MyArray(10)

Lower = Lbound(MyArray)     ' Returns 0 or 1, depending on setting of Option Base
```

### Redim 
Resize variables in [dynamic table](https://silkyroad.developpez.com/vba/tableaux/#LII-B). In order to keep the variables already present in the table, use the keyword *Preserve*.
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

## Error management

### IsError
Returns a boolean indicating whether an expression returns an error.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/iserror-function)
``` 
IsError(2/0) 'Return true
```

## Enumeration

### Enum
Enumeration variables are variables declared with an Enum type. Enum elements are initialized with constant values
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

## Manipulation of strings 

### Left
Returns a Variant (String) containing a specific number of characters on the left side of a string.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/left-function)
``` 
Dim myStr

myStr = Left("Hello World", 1)   ' Returns "H"
myStr = Left("Hello World", 7)   ' Returns "Hello W"
myStr = Left("Hello World", 20)  ' Returns "Hello World"
```

### Right
Returns a Variant value (String) containing a specified number of characters from the right end of a string.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/right-function)
``` 
Dim myStr

myStr = Right("Hello World", 1)    ' Returns "d"
myStr = Right("Hello World", 5)    ' Returns "World"
myStr = Right("Hello World", 20)   ' Returns "Hello World"
```

### Mid
Returns a Variant value (String) containing a specified number of characters extracted from a string.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/mid-function)
``` 
Dim myStr

myStr = Mid("Hello World", 1, 5)   ' Returns "Hello".
myStr = Mid("Hello World", 6, 5)   ' Returns "World".
myStr = Mid("Hello World", 5, 3)   ' Returns "o W".
```


### Split
Returns a one-dimensional zero-base array containing a specified number of substrings.
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
Returns a Variant (String) containing a copy of a specified string without leading spaces (LTrim), trailing spaces (RTrim), or both (Trim).
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/ltrim-rtrim-and-trim-functions)
``` 
Dim TrimString

TrimString = LTrim("  <-Trim->  ")    ' TrimString = "<-Trim->  "
TrimString = RTrim("  <-Trim->  ")    ' TrimString = "  <-Trim->"
TrimString = Trim("  <-Trim->  ")     ' TrimString = "<-Trim->"
```

## Class Module 

### class_Initialize and class_Terminate

[documentation sur les modules de classe](https://sinarf.developpez.com/access/vbaclass/#L2.4)


## Methodes

### Public
Public - Indicates that the procedure is accessible to all other procedures in all modules. If they are not explicitly specified with Public or Private **the procedures are public by default**
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/public-statement)
``` 
Public Number As Integer    ' Public Integer variable.
Public Property Get GetN() As Integer    ' Public Integer variable.
```

### Private
Private - Indicates that the procedure is only accessible to other procedures of the module in which it is declared
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/private-statement)
``` 
Private Number As Integer    ' Private Integer variable.
```

### ByVal
ByVal Indicates that the argument is passed by value.
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
ByRef Indicates that an argument is passed by reference. **The ByRef element is the default value in Visual Basic**.
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

Property **Get** allows reading a property of a class module
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/property-get-statement)
``` 
Private prvNom As String

Property Get Nom() As String
    ' Propriété en lecture
    Nom = prvNom
End Property
```


Property **Let** allows writing a property of a class module
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
Transfers control to a Sub procedure, a Function procedure
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/call-statement)
``` 
Call MySub
```

### Set 
Assign an object reference to a variable or property
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/set-statement)
``` 
Dim myRange As Range
Set myRange = Range("A1")
```
### New 
Create a new instance of an object.
[doc](https://stackoverflow.com/questions/21652671/what-does-the-keyword-new-do-in-vba)
``` 
Dim myObj As Object
Set myObj = New Object
```

### With
The With statement allows you to perform a series of instructions on a specified object without qualifying the object name each time.
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
Execute one or more groups of statements, depending on the value of an expression.
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
Repeats a group of instructions for each element in a Variant or Collection
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
Adds a named range to the workbook
[doc](https://docs.microsoft.com/en-gb/office/vba/api/excel.names.add)
``` 
ActiveWorkbook.Names.Add Name:="Game", RefersToR1C1:="=Feuil1!R2C2:R9C9"

Range("Game")   'refers to Range("B2:I9")
```

### Debug
Debug object can display values in the immediate window when debugging
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/debug-object)
``` 
Debug.Print "Hello Immediate window"
```
### Cstr
Convert to string
[doc](https://docs.microsoft.com/fr-fr/office/vba/language/reference/user-interface-help/returns-for-cstr)
``` 
Dim i as Integer 

i = 456789

Debug.Print Cstr(i)
```

### Timer
Returns the number of seconds elapsed since midnight.
[doc](https://docs.microsoft.com/en-gb/office/vba/language/reference/user-interface-help/timer-function)
``` 
Dim beginTimer
Dim endTimer

beginTimer = Timer

    'Do your code here

endTimer = Timer 
    
totalTimer = endTimer - beginTimer
Debug.Print Cstr(totalTimer)
```

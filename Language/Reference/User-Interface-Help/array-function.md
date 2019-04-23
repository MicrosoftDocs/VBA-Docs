---
title: Array function (Visual Basic for Applications)
keywords: vblr6.chm1010845
f1_keywords:
- vblr6.chm1010845
ms.prod: office
ms.assetid: dc7926a0-b70d-67ee-482f-d7bcdaffe139
ms.date: 12/11/2018
localization_priority: Priority
---


# Array function

Returns a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) containing an [array](../../Glossary/vbe-glossary.md#array).

## Syntax

**Array**(_arglist_)

The required _arglist_ [argument](../../Glossary/vbe-glossary.md#argument) is a comma-delimited list of values that are assigned to the elements of the array contained within the **Variant**. If no arguments are specified, an array of zero length is created.

## Remarks

The notation used to refer to an element of an array consists of the [variable](../../Glossary/vbe-glossary.md#variable) name followed by parentheses containing an index number indicating the desired element. 

In the following example, the first [statement](../../Glossary/vbe-glossary.md#statement) creates a variable named `A` as a **Variant**. The second statement assigns an array to variable `A`. The last statement assigns the value contained in the second array element to another variable.

```vb
Dim A As Variant, B As Long, i As Long
A = Array(10, 20, 30)  ' A is a three element list by default indexed 0 to 2
B = A(2)               ' B is now 30
ReDim Preserve A(4)    ' Extend A's length to five elements
A(4) = 40              ' Set the fifth element's value
For i = LBound(A) To UBound(A)
    Debug.Print "A(" & i & ") = " & A(i)
Next i

```

The lower bound of an array created by using the **Array** function is determined by the lower bound specified with the **[Option Base](option-base-statement.md)** statement, unless **Array** is qualified with the name of the type library (for example **VBA.Array**). If qualified with the type-library name, **Array** is unaffected by **Option Base**.

> [!NOTE] 
> A **Variant** that is not declared as an array can still contain an array. A **Variant** variable can contain an array of any type, except fixed-length strings and [user-defined types](../../Glossary/vbe-glossary.md#user-defined-type). Although a **Variant** containing an array is conceptually different from an array whose elements are of type **Variant**, the array elements are accessed in the same way.


## Example

This example uses the **Array** function to return a **Variant** containing an array.

```vb
Dim MyWeek, MyDay
MyWeek = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
' Return values assume lower bound set to 1 (using Option Base
' statement).
MyDay = MyWeek(2)    ' MyDay contains "Tue".
MyDay = MyWeek(4)    ' MyDay contains "Thu".
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

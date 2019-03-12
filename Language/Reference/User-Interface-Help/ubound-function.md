---
title: UBound function (Visual Basic for Applications)
keywords: vblr6.chm1009050
f1_keywords:
- vblr6.chm1009050
ms.prod: office
ms.assetid: 8dda22e9-d9f9-9944-1b91-cfb8b61774a7
ms.date: 12/13/2018
localization_priority: Normal
---


# UBound function

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) data type containing the largest available subscript for the indicated dimension of an [array](../../Glossary/vbe-glossary.md#array).

## Syntax

**UBound**(_arrayname_, [ _dimension_ ])

<br/>

The **UBound** function syntax has these parts.

|Part|Description|
|:-----|:-----|
| _arrayname_|Required. Name of the array [variable](../../Glossary/vbe-glossary.md#variable); follows standard variable naming conventions.|
| _dimension_|Optional; **Variant** (**Long**). Whole number indicating which dimension's upper bound is returned. Use 1 for the first dimension, 2 for the second, and so on. If _dimension_ is omitted, 1 is assumed.|

## Remarks

The **UBound** function is used with the **[LBound](lbound-function.md)** function to determine the size of an array. Use the **LBound** function to find the lower limit of an array dimension.

**UBound** returns the following values for an array with these dimensions:

|Statement|Return Value|
|:-----|:-----:|
| `UBound(A, 1)`|100|
| `UBound(A, 2)`|3|
| `UBound(A, 3)`|4|


## Example

This example uses the **UBound** function to determine the largest available subscript for the indicated dimension of an array.


```vb
Dim Upper
Dim MyArray(1 To 10, 5 To 15, 10 To 20)    ' Declare array variables.
Dim AnyArray(10)
Upper = UBound(MyArray, 1)    ' Returns 10.
Upper = UBound(MyArray, 3)    ' Returns 20.
Upper = UBound(AnyArray)      ' Returns 10.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

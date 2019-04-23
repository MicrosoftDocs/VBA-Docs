---
title: Erase statement (VBA)
keywords: vblr6.chm1008910
f1_keywords:
- vblr6.chm1008910
ms.prod: office
ms.assetid: b051ba13-3669-57e5-b023-cc4d52ec93f6
ms.date: 12/03/2018
localization_priority: Normal
---


# Erase statement

Reinitializes the elements of fixed-size [arrays](../../Glossary/vbe-glossary.md#array) and releases dynamic-array storage space.

## Syntax

**Erase** _arraylist_

The required _arraylist_ [argument](../../Glossary/vbe-glossary.md#argument) is one or more comma-delimited array [variables](../../Glossary/vbe-glossary.md#variable) to be erased.

## Remarks

**Erase** behaves differently depending on whether an array is fixed-size (ordinary) or dynamic. **Erase** recovers no memory for fixed-size arrays. **Erase** sets the elements of a fixed array as follows:


|Type of array|Effect of Erase on fixed-array elements|
|:-----|:-----|
|Fixed numeric array|Sets each element to zero.|
|Fixed string array (variable length)|Sets each element to a zero-length string ("").|
|Fixed string array (fixed length)|Sets each element to zero.|
|Fixed [Variant](../../Glossary/vbe-glossary.md#variant-data-type) array|Sets each element to [Empty](../../Glossary/vbe-glossary.md#empty).|
|Array of [user-defined types](../../Glossary/vbe-glossary.md#user-defined-type)|Sets each element as if it were a separate variable.|
|Array of objects|Sets each element to the special value **Nothing**.|

**Erase** frees the memory used by dynamic arrays. Before your program can refer to the dynamic array again, it must redeclare the array variable's dimensions by using a **[ReDim](redim-statement.md)** statement.

## Example

This example uses the **Erase** statement to reinitialize the elements of fixed-size arrays and deallocate dynamic-array storage space.


```vb
' Declare array variables. 
Dim NumArray(10) As Integer ' Integer array. 
Dim StrVarArray(10) As String ' Variable-string array. 
Dim StrFixArray(10) As String * 10 ' Fixed-string array. 
Dim VarArray(10) As Variant ' Variant array. 
Dim DynamicArray() As Integer ' Dynamic array. 
ReDim DynamicArray(10) ' Allocate storage space. 
Erase NumArray ' Each element set to 0. 
Erase StrVarArray ' Each element set to zero-length 
 ' string (""). 
Erase StrFixArray ' Each element set to 0. 
Erase VarArray ' Each element set to Empty. 
Erase DynamicArray ' Free memory used by array. 

```


## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

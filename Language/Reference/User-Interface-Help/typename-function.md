---
title: TypeName Function
keywords: vblr6.chm1010100
f1_keywords:
- vblr6.chm1010100
ms.prod: office
ms.assetid: 9353f1d5-5b64-9cad-5cc3-e1487bdd3afd
ms.date: 04/27/2018
---


# TypeName Function

Returns a  **String** that provides information about a [variable](../../Glossary/vbe-glossary.md#variable).</br></br>
## Syntax
**TypeName(**_varname_**)**</br>
The required _varname_ [argument](../../Glossary/vbe-glossary.md#argument) is a [Variant](../../Glossary/vbe-glossary.md#Variant) containing any variable except a variable of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type).
## Remarks
The string returned by **TypeName** can be any one of the following:</br>

|**String returned**|**Variable**|
|:-----|:-----|
|[object type](../../Glossary/vbe-glossary.md#object-type)|An object whose type is  _objecttype_|
|[Byte](../../Glossary/vbe-glossary.md#Byte)|Byte value|
|[Integer](../../Glossary/vbe-glossary.md#Integer)|Integer|
|[Long](../../Glossary/vbe-glossary.md#Long)|Long integer|
|[Single](../../Glossary/vbe-glossary.md#Single)|Single-precision floating-point number|
|[Double](../../Glossary/vbe-glossary.md#Double)|Double-precision floating-point number|
|[Currency](../../Glossary/vbe-glossary.md#Currency)|Currency value|
|[Decimal](../../Glossary/vbe-glossary.md#Decimal)|Decimal value|
|[Date](../../Glossary/vbe-glossary.md#Date)|Date value|
|[String](../../Glossary/vbe-glossary.md#String)|String|
|[Boolean](../../Glossary/vbe-glossary.md#Boolean)|Boolean value|
|**Error**|An error value|
|[Empty](../../Glossary/vbe-glossary.md#Empty)|Uninitialized|
|[Null](../../Glossary/vbe-glossary.md#Null)|No valid data|
|[Object](../../Glossary/vbe-glossary.md#Object)|An object|
|Unknown|An object whose type is unknown|
|**Nothing**|Object variable that doesn't refer to an object|

<br>
If  _varname_ is an [array](../../Glossary/vbe-glossary.md#array), the returned string can be any one of the possible returned strings (or  **Variant**) with empty parentheses appended. For example, if _varname_ is an array of integers, **TypeName** returns `"Integer()`".

## Example

This example uses the **TypeName** function to return information about a variable.


```vb
' Declare variables.
Dim NullVar, MyType, StrVar As String, IntVar As Integer, CurVar As Currency
Dim ArrayVar (1 To 5) As Integer
NullVar = Null    ' Assign Null value.
MyType = TypeName(StrVar)    ' Returns "String".
MyType = TypeName(IntVar)    ' Returns "Integer".
MyType = TypeName(CurVar)    ' Returns "Currency".
MyType = TypeName(NullVar)    ' Returns "Null".
MyType = TypeName(ArrayVar)    ' Returns "Integer()".

```



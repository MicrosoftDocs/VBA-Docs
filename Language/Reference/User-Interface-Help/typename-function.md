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

<<<<<<< HEAD
Returns a  **String** that provides information about a [variable](../../Glossary/vbe-glossary.md).</br></br>
## Syntax
**TypeName(**_varname_**)**</br>
The required _varname_ [argument](../../Glossary/vbe-glossary.md) is a [Variant](../../Glossary/vbe-glossary.md) containing any variable except a variable of a [user-defined type](../../Glossary/vbe-glossary.md).
=======
Returns a  **String** that provides information about a [variable](../../Glossary/vbe-glossary.md#variable).</br></br>
## Syntax
**TypeName(**_varname_**)**</br>
The required _varname_ [argument](../../Glossary/vbe-glossary.md#argument) is a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) containing any variable except a variable of a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type).
>>>>>>> master
## Remarks
The string returned by **TypeName** can be any one of the following:</br>

|**String returned**|**Variable**|
|:-----|:-----|
<<<<<<< HEAD
|[object type](../../Glossary/vbe-glossary.md)|An object whose type is  _objecttype_|
|[Byte](../../Glossary/vbe-glossary.md)|Byte value|
|[Integer](../../Glossary/vbe-glossary.md)|Integer|
|[Long](../../Glossary/vbe-glossary.md)|Long integer|
|[Single](../../Glossary/vbe-glossary.md)|Single-precision floating-point number|
|[Double](../../Glossary/vbe-glossary.md)|Double-precision floating-point number|
|[Currency](../../Glossary/vbe-glossary.md)|Currency value|
|[Decimal](../../Glossary/vbe-glossary.md)|Decimal value|
|[Date](../../Glossary/vbe-glossary.md)|Date value|
|[String](../../Glossary/vbe-glossary.md)|String|
|[Boolean](../../Glossary/vbe-glossary.md)|Boolean value|
|**Error**|An error value|
|[Empty](../../Glossary/vbe-glossary.md)|Uninitialized|
|[Null](../../Glossary/vbe-glossary.md)|No valid data|
|[Object](../../Glossary/vbe-glossary.md)|An object|
=======
|[object type](../../Glossary/vbe-glossary.md#object-type)|An object whose type is  _objecttype_|
|[Byte](../../Glossary/vbe-glossary.md#byte-data-type)|Byte value|
|[Integer](../../Glossary/vbe-glossary.md#integer-data-type)|Integer|
|[Long](../../Glossary/vbe-glossary.md#long-data-type)|Long integer|
|[Single](../../Glossary/vbe-glossary.md#single-data-type)|Single-precision floating-point number|
|[Double](../../Glossary/vbe-glossary.md#double-data-type)|Double-precision floating-point number|
|[Currency](../../Glossary/vbe-glossary.md#currency-data-type)|Currency value|
|[Decimal](../../Glossary/vbe-glossary.md#decimal-data-type)|Decimal value|
|[Date](../../Glossary/vbe-glossary.md#date-data-type)|Date value|
|[String](../../Glossary/vbe-glossary.md#string-data-type)|String|
|[Boolean](../../Glossary/vbe-glossary.md#boolean-data-type)|Boolean value|
|**Error**|An error value|
|[Empty](../../Glossary/vbe-glossary.md#empty)|Uninitialized|
|[Null](../../Glossary/vbe-glossary.md#null)|No valid data|
|[Object](../../Glossary/vbe-glossary.md#object)|An object|
>>>>>>> master
|Unknown|An object whose type is unknown|
|**Nothing**|Object variable that doesn't refer to an object|

<br>
<<<<<<< HEAD
If  _varname_ is an [array](../../Glossary/vbe-glossary.md), the returned string can be any one of the possible returned strings (or  **Variant**) with empty parentheses appended. For example, if _varname_ is an array of integers, **TypeName** returns `"Integer()`".
=======
If  _varname_ is an [array](../../Glossary/vbe-glossary.md#array), the returned string can be any one of the possible returned strings (or  **Variant**) with empty parentheses appended. For example, if _varname_ is an array of integers, **TypeName** returns `"Integer()`".
>>>>>>> master

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



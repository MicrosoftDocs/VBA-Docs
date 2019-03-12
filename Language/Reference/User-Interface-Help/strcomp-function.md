---
title: StrComp function (Visual Basic for Applications)
keywords: vblr6.chm1009035
f1_keywords:
- vblr6.chm1009035
ms.prod: office
ms.assetid: 96b0e82d-e080-0c60-94d1-ddff10d6ee86
ms.date: 12/13/2018
localization_priority: Normal
---


# StrComp function

Returns a **Variant** (**Integer**) indicating the result of a [string comparison](../../Glossary/vbe-glossary.md#string-comparison).

## Syntax

**StrComp**(_string1_, _string2_, [ _compare_ ])

<br/>

The **StrComp** function syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_string1_|Required. Any valid [string expression](../../Glossary/vbe-glossary.md#string-expression).|
|_string2_|Required. Any valid string expression.|
|_compare_|Optional. Specifies the type of string comparison. If the _compare_ [argument](../../Glossary/vbe-glossary.md#argument) is [Null](../../Glossary/vbe-glossary.md#null), an error occurs. If _compare_ is omitted, the **Option Compare** setting determines the type of comparison.|

## Settings

The _compare_ argument settings are:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbUseCompareOption**|-1|Performs a comparison by using the setting of the **[Option Compare](option-compare-statement.md)** statement.|
|**vbBinaryCompare**|0|Performs a binary comparison.|
|**vbTextCompare**|1|Performs a textual comparison.|
|**vbDatabaseCompare**|2|Microsoft Access only. Performs a comparison based on information in your database.|

## Return values

The **StrComp** function has the following return values:

|If|StrComp returns|
|:-----|:-----|
|_string1_ is less than _string2_|-1|
|_string1_ is equal to _string2_|0|
|_string1_ is greater than _string2_|1|
|_string1_ or _string2_ is **Null**|**Null**|

## Example

This example uses the **StrComp** function to return the results of a string comparison. If the third argument is 1, a textual comparison is performed; if the third argument is 0 or omitted, a binary comparison is performed.

```vb
Dim MyStr1, MyStr2, MyComp
MyStr1 = "ABCD": MyStr2 = "abcd"    ' Define variables.
MyComp = StrComp(MyStr1, MyStr2, 1)    ' Returns 0.
MyComp = StrComp(MyStr1, MyStr2, 0)    ' Returns -1.
MyComp = StrComp(MyStr2, MyStr1)    ' Returns 1.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

---
title: Filter function (Visual Basic for Applications)
keywords: vblr6.chm1008912
f1_keywords:
- vblr6.chm1008912
ms.prod: office
ms.assetid: 00630b25-e7b8-5c32-b6d1-9816f01c3a0f
ms.date: 12/12/2018
localization_priority: Normal
---


# Filter function

Returns a zero-based array containing a subset of a string array based on a specified filter criteria.

## Syntax

**Filter**(_sourcearray_, _match_, [ _include_, [ _compare_ ]])

<br/>

The **Filter** function syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_sourcearray_|Required. One-dimensional array of strings to be searched.|
|_match_|Required. String to search for.|
|_include_|Optional. **Boolean** value indicating whether to return substrings that include or exclude _match_. If _include_ is **True**, **Filter** returns the subset of the array that contains _match_ as a substring. If _include_ is **False**, **Filter** returns the subset of the array that does not contain _match_ as a substring.|
|_compare_|Optional. Numeric value indicating the kind of string comparison to use. See Settings section for values.|

## Settings

The _compare_ argument can have the following values:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbUseCompareOption**|-1|Performs a comparison by using the setting of the **[Option Compare](option-compare-statement.md)** statement.|
|**vbBinaryCompare**| 0|Performs a binary comparison.|
|**vbTextCompare**| 1|Performs a textual comparison.|
|**vbDatabaseCompare**| 2|Microsoft Access only. Performs a comparison based on information in your database.|

The array returned by the **Filter** function contains only enough elements to contain the number of matched items.

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

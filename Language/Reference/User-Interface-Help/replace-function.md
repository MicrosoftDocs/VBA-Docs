---
title: Replace function (Visual Basic for Applications)
keywords: vblr6.chm1008930
f1_keywords:
- vblr6.chm1008930
ms.prod: office
ms.assetid: a24e3da4-fc94-56e7-d718-f4c2d0a31072
ms.date: 12/13/2018
localization_priority: Normal
---


# Replace function

Returns a string, which is a substring of a string expression beginning at the start position (defaults to 1), in which a specified substring has been replaced with another substring a specified number of times.

## Syntax

**Replace**(_expression_, _find_, _replace_, [ _start_, [ _count_, [ _compare_ ]]])

<br/>

The **Replace** function syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_expression_|Required. [String expression](../../Glossary/vbe-glossary.md#string-expression) containing substring to replace.|
|_find_|Required. Substring being searched for.|
|_replace_|Required. Replacement substring.|
|_start_|Optional. Start position for the substring of _expression_ to be searched and returned. If omitted, 1 is assumed.|
|_count_|Optional. Number of substring substitutions to perform. If omitted, the default value is -1, which means, make all possible substitutions.|
|_compare_|Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. See Settings section for values.|

## Settings

The _compare_ argument can have the following values:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbUseCompareOption**|-1|Performs a comparison by using the setting of the **[Option Compare](option-compare-statement.md)** statement.|
|**vbBinaryCompare**|0|Performs a binary comparison.|
|**vbTextCompare**|1|Performs a textual comparison.|
|**vbDatabaseCompare**|2|Microsoft Access only. Performs a comparison based on information in your database.|

## Return values

**Replace** returns the following values:

|If|Replace returns|
|:-----|:-----|
|_expression_ is zero-length|Zero-length string ("")|
|_expression_ is **Null**|An error.|
|_find_ is zero-length|Copy of _expression_.|
|_replace_ is zero-length|Copy of _expression_ with all occurrences of _find_ removed.|
|_start_ > **Len**(_expression_)|Zero-length string. String replacement begins at the position indicated by _start_.|
|_count_ is 0|Copy of _expression_.|

## Remarks

The return value of the **Replace** function is a string, with substitutions made, that begins at the position specified by _start_ and concludes at the end of the _expression_ string. It is not a copy of the original string from start to finish.

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

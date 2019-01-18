---
title: CompareMode property (Visual Basic for Applications)
keywords: vblr6.chm2181931
f1_keywords:
- vblr6.chm2181931
ms.prod: office
api_name:
- Office.CompareMode
ms.assetid: 75893886-8bed-4685-b483-18b3d39569da
ms.date: 12/19/2018
localization_priority: Normal
---


# CompareMode property

Sets and returns the comparison mode for comparing string keys in a **[Dictionary](dictionary-object.md)** object.

## Syntax

_object_.**CompareMode** [ = _compare_ ]

<br/>

The **CompareMode** property has the following parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **Dictionary** object.|
| _compare_|Optional. If provided, _compare_ is a value representing the comparison mode used by functions such as **[StrComp](strcomp-function.md)**.|

## Settings

The _compare_ argument can have the following values:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbUseCompareOption**|-1|Performs a comparison by using the setting of the **[Option Compare](option-compare-statement.md)** statement.|
|**vbBinaryCompare**| 0|Performs a binary comparison.|
|**vbTextCompare**| 1|Performs a textual comparison.|
|**vbDatabaseCompare**| 2|Microsoft Access only. Performs a comparison based on information in your database.|

## Remarks

An error occurs if you try to change the comparison mode of a **Dictionary** object that already contains data.

The **CompareMode** property uses the same values as the _compare_ argument for the **StrComp** function. Values greater than 2 can be used to refer to comparisons by using specific Locale IDs (LCID).

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
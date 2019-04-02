---
title: ChartTitle.Characters property (PowerPoint)
keywords: vbapp10.chm694002
f1_keywords:
- vbapp10.chm694002
ms.prod: powerpoint
api_name:
- PowerPoint.ChartTitle.Characters
ms.assetid: 38d03ba1-54dd-4bcb-1da3-37ad9b4f9b4c
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartTitle.Characters property (PowerPoint)

Returns a  **[ChartCharacters](PowerPoint.ChartCharacters.md)** object that represents a range of characters within the object text. You can use the **ChartCharacters** object to format characters within a text string.


## Syntax

_expression_. `Characters`( `_Start_`, `_Length_` )

_expression_ A variable that represents a '[ChartTitle](PowerPoint.ChartTitle.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional|**Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the Start character).|

## Remarks

The  **ChartCharacters** object is not a collection.


## See also


[ChartTitle Object](PowerPoint.ChartTitle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
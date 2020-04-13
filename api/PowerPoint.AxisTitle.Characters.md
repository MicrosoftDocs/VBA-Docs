---
title: AxisTitle.Characters property (PowerPoint)
keywords: vbapp10.chm683002
f1_keywords:
- vbapp10.chm683002
ms.prod: powerpoint
api_name:
- PowerPoint.AxisTitle.Characters
ms.assetid: 8b1b9dc9-6aa3-872f-964a-fe623feff6fa
ms.date: 06/08/2017
localization_priority: Normal
---


# AxisTitle.Characters property (PowerPoint)

Returns a **[ChartCharacters](PowerPoint.ChartCharacters.md)** object that represents a range of characters within the object text. You can use the **ChartCharacters** object to format characters within a text string.


## Syntax

_expression_. `Characters`( `_Start_`, `_Length_` )

_expression_ A variable that represents an '[AxisTitle](PowerPoint.AxisTitle.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional|**Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional|**Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the Start character).|

## Remarks

The **ChartCharacters** object is not a collection.


## See also


[AxisTitle Object](PowerPoint.AxisTitle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
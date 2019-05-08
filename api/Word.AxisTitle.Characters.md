---
title: AxisTitle.Characters property (Word)
keywords: vbawd10.chm98238466
f1_keywords:
- vbawd10.chm98238466
ms.prod: word
api_name:
- Word.AxisTitle.Characters
ms.assetid: a53ea4c6-a742-17b0-8120-1c16484c5995
ms.date: 06/08/2017
localization_priority: Normal
---


# AxisTitle.Characters property (Word)

Returns a  **[ChartCharacters](Word.ChartCharacters.md)** object that represents a range of characters within the object text. You can use the **ChartCharacters** object to format characters within a text string.


## Syntax

_expression_.**Characters** (_Start_, _Length_)

_expression_ A variable that represents an '[AxisTitle](Word.AxisTitle.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional| **Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the Start character).|

## Remarks

The  **ChartCharacters** object is not a collection.


## See also


[AxisTitle Object](Word.AxisTitle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
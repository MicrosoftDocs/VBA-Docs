---
title: ChartTitle.Characters property (Word)
keywords: vbawd10.chm65273858
f1_keywords:
- vbawd10.chm65273858
ms.prod: word
api_name:
- Word.ChartTitle.Characters
ms.assetid: 24650d31-1618-b231-ce3e-d7f35f39db5b
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartTitle.Characters property (Word)

Returns a  **[ChartCharacters](Word.ChartCharacters.md)** object that represents a range of characters within the object text. You can use the **ChartCharacters** object to format characters within a text string.


## Syntax

_expression_.**Characters** (_Start_, _Length_)

_expression_ A variable that represents a '[ChartTitle](Word.ChartTitle.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional| **Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the Start character).|

## Remarks

The  **ChartCharacters** object is not a collection.


## See also


[ChartTitle Object](Word.ChartTitle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
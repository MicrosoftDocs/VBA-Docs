---
title: Range.Speak method (Excel)
keywords: vbaxl10.chm144237
f1_keywords:
- vbaxl10.chm144237
api_name:
- Excel.Range.Speak
ms.assetid: 12928814-9534-c9f0-e351-7d26f77869e0
ms.date: 05/11/2019
ms.localizationpriority: medium
---


# Range.Speak method (Excel)

Causes the cells of the range to be spoken in row order or column order.


## Syntax

_expression_.**Speak** (_SpeakDirection_, _SpeakFormulas_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SpeakDirection_|Optional| **Variant**|The speak direction, by rows or columns.|
| _SpeakFormulas_|Optional| **Variant**| **True** will cause formulas to be sent to the Text-To-Speech (TTS) engine for cells that have formulas. The value is sent if the cells don't have formulas. **False** (default) will cause values to always be sent to the TTS engine.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
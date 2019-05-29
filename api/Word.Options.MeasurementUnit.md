---
title: Options.MeasurementUnit property (Word)
keywords: vbawd10.chm162988058
f1_keywords:
- vbawd10.chm162988058
ms.prod: word
api_name:
- Word.Options.MeasurementUnit
ms.assetid: 7d5b1c89-eedd-9818-2137-d94e6f80d629
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MeasurementUnit property (Word)

Returns or sets the standard measurement unit for Microsoft Word. Read/write  **WdMeasurementUnits**.


## Syntax

_expression_. `MeasurementUnit`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets the standard measurement unit for Word to points.


```vb
Options.MeasurementUnit = wdPoints
```

This example returns the current measurement unit selected on the General tab in the Options dialog box (Tools menu).




```vb
CurrUnit = Options.MeasurementUnit
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
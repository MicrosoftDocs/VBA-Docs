---
title: DataLabel.AutoText property (Excel)
keywords: vbaxl10.chm582092
f1_keywords:
- vbaxl10.chm582092
ms.prod: excel
api_name:
- Excel.DataLabel.AutoText
ms.assetid: a549b738-59fb-a096-c4e9-d8f00bc59239
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.AutoText property (Excel)

 **True** if the object automatically generates appropriate text based on context. Read/write **Boolean**.


## Syntax

_expression_. `AutoText`

_expression_ A variable that represents a [DataLabel](Excel.DataLabel-graph-property.md) object.


## Example

This example sets the data labels for series one in Chart1 to automatically generate appropriate text.


```vb
Charts("Chart1").SeriesCollection(1).DataLabels.AutoText = True
```


## See also


[DataLabel Object](Excel.DataLabel(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
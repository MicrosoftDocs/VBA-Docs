---
title: Series.ApplyPictToSides property (Excel)
keywords: vbaxl10.chm578115
f1_keywords:
- vbaxl10.chm578115
ms.prod: excel
api_name:
- Excel.Series.ApplyPictToSides
ms.assetid: 300e9c75-4108-32bc-01ab-c622843e6fbd
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.ApplyPictToSides property (Excel)

 **True** if a picture is applied to the sides of the point or all points in the series. Read/write **Boolean**.


## Syntax

_expression_. `ApplyPictToSides`

_expression_ A variable that represents a [Series](Excel.Series-graph-object.md) object.


## Example

This example applies pictures to the sides of all points in series one. The series must already have pictures applied to it (setting this property changes the picture orientation).


```vb
Charts(1).SeriesCollection(1).ApplyPictToSides = True
```


## See also


[Series Object](Excel.Series(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
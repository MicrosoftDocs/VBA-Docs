---
title: Series.PictureUnit2 property (Excel)
keywords: vbaxl10.chm578123
f1_keywords:
- vbaxl10.chm578123
ms.prod: excel
api_name:
- Excel.Series.PictureUnit2
ms.assetid: 6c29fd60-2e42-3f7a-1fc0-67309ea75a3a
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.PictureUnit2 property (Excel)

Returns or sets the unit for each picture on the chart if the **[PictureType](Excel.Series.PictureType.md)** property is set to **xlStackScale** (if not, this property is ignored). Read/write **Double**.


## Syntax

_expression_.**PictureUnit2**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Example

This example sets series one on Chart1 to stack pictures, and uses each picture to represent five units. The example should be run on a 2D column chart with picture data markers.

```vb
With Charts("Chart1").SeriesCollection(1) 
 .PictureType = xlScale 
 .PictureUnit2 = 5 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
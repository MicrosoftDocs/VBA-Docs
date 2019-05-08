---
title: Point.Paste method (Excel)
keywords: vbaxl10.chm576090
f1_keywords:
- vbaxl10.chm576090
ms.prod: excel
api_name:
- Excel.Point.Paste
ms.assetid: 0a984f1c-54de-d49f-8677-43d513a0f9fc
ms.date: 05/09/2019
localization_priority: Normal
---


# Point.Paste method (Excel)

Pastes a picture from the Clipboard as the marker on the selected point.


## Syntax

_expression_.**Paste**

_expression_ A variable that represents a **[Point](Excel.Point(object).md)** object.


## Return value

Variant


## Remarks

This method can be used on column, bar, line, or radar charts, and it sets the **[MarkerStyle](Excel.Point.MarkerStyle.md)** property to **xlMarkerStylePicture**.


## Example

This example pastes a picture from the Clipboard into series one on Chart1.

```vb
Charts("Chart1").SeriesCollection(1).Paste
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
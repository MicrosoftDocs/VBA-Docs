---
title: Series.Paste method (Excel)
keywords: vbaxl10.chm578100
f1_keywords:
- vbaxl10.chm578100
ms.prod: excel
api_name:
- Excel.Series.Paste
ms.assetid: 73e689cb-b2aa-61d7-e84c-113091d09a44
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.Paste method (Excel)

Pastes a picture from the Clipboard as the marker on the selected series.


## Syntax

_expression_.**Paste**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Return value

Variant


## Remarks

This method can be used on column, bar, line, or radar charts, and it sets the **[MarkerStyle](Excel.Series.MarkerStyle.md)** property to **xlMarkerStylePicture**.


## Example

This example pastes a picture from the Clipboard into series one on Chart1.

```vb
Charts("Chart1").SeriesCollection(1).Paste
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
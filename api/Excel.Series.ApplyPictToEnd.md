---
title: Series.ApplyPictToEnd property (Excel)
keywords: vbaxl10.chm578117
f1_keywords:
- vbaxl10.chm578117
ms.prod: excel
api_name:
- Excel.Series.ApplyPictToEnd
ms.assetid: 40d4dca5-1747-c9c6-a117-29939bf4cd74
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.ApplyPictToEnd property (Excel)

**True** if a picture is applied to the end of the point or all points in the series. Read/write **Boolean**.


## Syntax

_expression_.**ApplyPictToEnd**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Example

This example applies pictures to the end of all points in series one. The series must already have pictures applied to it (setting this property changes the picture orientation).

```vb
Charts(1).SeriesCollection(1).ApplyPictToEnd = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
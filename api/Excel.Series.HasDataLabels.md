---
title: Series.HasDataLabels property (Excel)
keywords: vbaxl10.chm578088
f1_keywords:
- vbaxl10.chm578088
ms.prod: excel
api_name:
- Excel.Series.HasDataLabels
ms.assetid: 10f879c9-4d34-d20b-facc-44ebc950aaa2
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.HasDataLabels property (Excel)

**True** if the series has data labels. Read/write **Boolean**.


## Syntax

_expression_.**HasDataLabels**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Example

This example turns on data labels for series three on Chart1.

```vb
With Charts("Chart1").SeriesCollection(3) 
 .HasDataLabels = True 
 .ApplyDataLabels Type:=xlValue 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
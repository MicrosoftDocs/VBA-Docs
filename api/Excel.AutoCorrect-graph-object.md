---
title: AutoCorrect object (Excel Graph)
keywords: vbagr10.chm3077641
f1_keywords:
- vbagr10.chm3077641
ms.prod: excel
api_name:
- Excel.AutoCorrect
ms.assetid: 68fa11da-e28f-53cd-3d50-a1f19d261a02
ms.date: 04/06/2019
localization_priority: Normal
---


# AutoCorrect object (Excel Graph)

Contains Graph AutoCorrect attributes (capitalization of names of days, correction of two initial capital letters, automatic correction list, and so on).

## Remarks

Use the **[AutoCorrect](Excel.AutoCorrect-graph-property.md)** property to return the **AutoCorrect** object. 

## Example

The following example sets Graph to correct words that begin with two initial capital letters.

```vb
With myChart.Application.AutoCorrect 
 .TwoInitialCapitals = True 
 .ReplaceText = True 
End With
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
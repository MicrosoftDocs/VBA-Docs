---
title: AutoCorrect object (Excel)
keywords: vbaxl10.chm544072
f1_keywords:
- vbaxl10.chm544072
ms.prod: excel
api_name:
- Excel.AutoCorrect
ms.assetid: 2594722a-2ff9-7175-4d35-0da0ad413b0d
ms.date: 03/29/2019
localization_priority: Normal
---


# AutoCorrect object (Excel)

Contains Microsoft Excel AutoCorrect attributes (capitalization of names of days, correction of two initial capital letters, automatic correction list, and so on).


## Example

Use the **[AutoCorrect](Excel.Application.AutoCorrect.md)** property of the **Application** object to return the **AutoCorrect** object. The following example sets Excel to correct words that begin with two initial capital letters.

```vb
With Application.AutoCorrect 
 .TwoInitialCapitals = True 
 .ReplaceText = True 
End With
```


## Methods

- [AddReplacement](Excel.AutoCorrect.AddReplacement.md)
- [DeleteReplacement](Excel.AutoCorrect.DeleteReplacement.md)

## Properties

- [Application](Excel.AutoCorrect.Application.md)
- [AutoExpandListRange](Excel.AutoCorrect.AutoExpandListRange.md)
- [AutoFillFormulasInLists](Excel.AutoCorrect.AutoFillFormulasInLists.md)
- [CapitalizeNamesOfDays](Excel.AutoCorrect.CapitalizeNamesOfDays.md)
- [CorrectCapsLock](Excel.AutoCorrect.CorrectCapsLock.md)
- [CorrectSentenceCap](Excel.AutoCorrect.CorrectSentenceCap.md)
- [Creator](Excel.AutoCorrect.Creator.md)
- [DisplayAutoCorrectOptions](Excel.AutoCorrect.DisplayAutoCorrectOptions.md)
- [Parent](Excel.AutoCorrect.Parent.md)
- [ReplacementList](Excel.AutoCorrect.ReplacementList.md)
- [ReplaceText](Excel.AutoCorrect.ReplaceText.md)
- [TwoInitialCapitals](Excel.AutoCorrect.TwoInitialCapitals.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
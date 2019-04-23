---
title: SpellingOptions object (Excel)
keywords: vbaxl10.chm717072
f1_keywords:
- vbaxl10.chm717072
ms.prod: excel
api_name:
- Excel.SpellingOptions
ms.assetid: 3ba7d0b4-bebb-0cc9-cb50-066d1c19d876
ms.date: 04/02/2019
localization_priority: Normal
---


# SpellingOptions object (Excel)

Represents the various spell checking options for a worksheet.


## Remarks

Use the **[SpellingOptions](Excel.Application.SpellingOptions.md)** property of the **Application** object to return a **SpellingOptions** object.

After a **SpellingOptions** object is returned, you can use the following properties to set or return various spell checking options:

- **ArabicModes**   
- **DictLang**    
- **GermanPostReform**   
- **HebrewModes**   
- **IgnoreCaps**    
- **IgnoreFileNames**   
- **IgnoreMixedDigits**   
- **KoreanCombineAux**    
- **KoreanProcessCompound**    
- **KoreanUseAutoChangeList**    
- **SuggestMainOnly**    
- **UserDict**
    

## Example

The following example uses the **IgnoreCaps** property to disable spell checking for words that have all capitalized letters. In this example, "Testt", but not "TESTT", is identified by the spell checker.

```vb
Sub IgnoreAllCAPS() 
 
 ' Place misspelled versions of the same word in all caps and mixed case. 
 Range("A1").Formula = "Testt" 
 Range("A2").Formula = "TESTT" 
 
 With Application.SpellingOptions 
 .SuggestMainOnly = True 
 .IgnoreCaps = True 
 End With 
 
 ' Run a spell check. 
 Cells.CheckSpelling 
 
End Sub
```

## Properties

- [ArabicModes](Excel.SpellingOptions.ArabicModes.md)
- [ArabicStrictAlefHamza](Excel.SpellingOptions.ArabicStrictAlefHamza.md)
- [ArabicStrictFinalYaa](Excel.SpellingOptions.ArabicStrictFinalYaa.md)
- [ArabicStrictTaaMarboota](Excel.SpellingOptions.ArabicStrictTaaMarboota.md)
- [BrazilReform](Excel.SpellingOptions.BrazilReform.md)
- [DictLang](Excel.SpellingOptions.DictLang.md)
- [GermanPostReform](Excel.SpellingOptions.GermanPostReform.md)
- [HebrewModes](Excel.SpellingOptions.HebrewModes.md)
- [IgnoreCaps](Excel.SpellingOptions.IgnoreCaps.md)
- [IgnoreFileNames](Excel.SpellingOptions.IgnoreFileNames.md)
- [IgnoreMixedDigits](Excel.SpellingOptions.IgnoreMixedDigits.md)
- [KoreanCombineAux](Excel.SpellingOptions.KoreanCombineAux.md)
- [KoreanProcessCompound](Excel.SpellingOptions.KoreanProcessCompound.md)
- [KoreanUseAutoChangeList](Excel.SpellingOptions.KoreanUseAutoChangeList.md)
- [PortugalReform](Excel.SpellingOptions.PortugalReform.md)
- [RussianStrictE](Excel.SpellingOptions.RussianStrictE.md)
- [SpanishModes](Excel.SpellingOptions.SpanishModes.md)
- [SuggestMainOnly](Excel.SpellingOptions.SuggestMainOnly.md)
- [UserDict](Excel.SpellingOptions.UserDict.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
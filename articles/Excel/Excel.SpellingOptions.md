---
title: SpellingOptions Object (Excel)
keywords: vbaxl10.chm717072
f1_keywords:
- vbaxl10.chm717072
ms.prod: excel
api_name:
- Excel.SpellingOptions
ms.assetid: 3ba7d0b4-bebb-0cc9-cb50-066d1c19d876
ms.date: 06/08/2017
---


# SpellingOptions Object (Excel)

Represents the various spell checking options for a worksheet.


## Remarks

Use the  **[SpellingOptions](Excel.Application.SpellingOptions.md)** property of the **[Application](Excel.Application(objec).md)** object to return a **SpellingOptions** object.

Once a  **SpellingOptions** object is returned, you can use its following properties to set or return various spell checking options.


-  **[ArabicModes](Excel.SpellingOptions.ArabicModes.md)**
    
-  **[DictLang](Excel.SpellingOptions.DictLang.md)**
    
-  **[GermanPostReform](Excel.SpellingOptions.GermanPostReform.md)**
    
-  **[HebrewModes](Excel.SpellingOptions.HebrewModes.md)**
    
-  **[IgnoreCaps](Excel.SpellingOptions.IgnoreCaps.md)**
    
-  **[IgnoreFileNames](Excel.SpellingOptions.IgnoreFileNames.md)**
    
-  **[IgnoreMixedDigits](Excel.SpellingOptions.IgnoreMixedDigits.md)**
    
-  **[KoreanCombineAux](Excel.SpellingOptions.KoreanCombineAux.md)**
    
-  **[KoreanProcessCompound](Excel.SpellingOptions.KoreanProcessCompound.md)**
    
-  **[KoreanUseAutoChangeList](Excel.SpellingOptions.KoreanUseAutoChangeList.md)**
    
-  **[SuggestMainOnly](Excel.SpellingOptions.SuggestMainOnly.md)**
    
-  **[UserDict](Excel.SpellingOptions.UserDict.md)**
    

## Example

The following example uses the  **[IgnoreCaps](Excel.SpellingOptions.IgnoreCaps.md)** property to disable spell checking for words that have all capitalized letters. In this example, "Testt", but not "TESTT", is identified by the spell checker.


```vb
Sub IgnoreAllCAPS() 
 
 ' Place mispelled versions of the same word in all caps and mixed case. 
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


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)



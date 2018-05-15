---
title: FindReplace Object (Publisher)
keywords: vbapb10.chm8388607
f1_keywords:
- vbapb10.chm8388607
ms.prod: publisher
api_name:
- Publisher.FindReplace
ms.assetid: 96dcf5fe-4f3e-07b7-c248-46edd370fc31
ms.date: 06/08/2017
---


# FindReplace Object (Publisher)

Represents the criteria for a find operation. The properties and methods of the  **FindReplace** object correspond to the options in the **Find and Replace** dialog box.
 


## Remarks

When the  **ReplaceScope** property is set to **pbReplaceScopeOne** or **pbReplaceScopeAll**, the **ReplaceWithText** property must be set to avoid the text from being replaced with the default value of an empty **String** for that property.
 

 

## Example

Use the  **Find** property to return a **FindReplace** object. The following example selects the next occurrence of the word "factory".
 

 

```
With ActiveDocument.Find 
 .Clear 
 .FindText = "factory" 
 .Execute 
End With
```

Set the  **ReplaceScope** property to determine the extent of the search. The following example replaces the first occurrence of the name "Visual Basic Scripting Edition" with "VBScript".
 

 



```
With ActiveDocument.Find 
 .Clear 
 .FindText = "Visual Basic Scripting Edition" 
 .ReplaceWithText = "VBScript" 
 .ReplaceScope = pbReplaceScopeOne 
 .Execute 
End With
```

The following example illustrates how the font attributes of the FoundTextRange can be accessed when  **ReplaceScope** is set to **pbReplaceScopeNone**.
 

 



```
Dim objFindReplace As FindReplace 
 
Set objFindReplace = ActiveDocument.Find 
With objFindReplace 
 .Clear 
 .FindText = "important" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 If .FoundTextRange.Font.Italic = msoFalse Then 
 .FoundTextRange.Font.Italic = msoTrue 
 End If 
 Loop 
End With
```


## Methods



|**Name**|
|:-----|
|[Clear](Publisher.FindReplace.Clear.md)|
|[Execute](Publisher.FindReplace.Execute.md)|

## Properties



|**Name**|
|:-----|
|[Application](Publisher.FindReplace.Application.md)|
|[FindText](Publisher.FindReplace.FindText.md)|
|[Forward](Publisher.FindReplace.Forward.md)|
|[FoundTextRange](Publisher.FindReplace.FoundTextRange.md)|
|[MatchAlefHamza](Publisher.FindReplace.MatchAlefHamza.md)|
|[MatchCase](Publisher.FindReplace.MatchCase.md)|
|[MatchDiacritics](Publisher.FindReplace.MatchDiacritics.md)|
|[MatchKashida](Publisher.FindReplace.MatchKashida.md)|
|[MatchWholeWord](Publisher.FindReplace.MatchWholeWord.md)|
|[MatchWidth](Publisher.FindReplace.MatchWidth.md)|
|[Parent](Publisher.FindReplace.Parent.md)|
|[ReplaceScope](Publisher.FindReplace.ReplaceScope.md)|
|[ReplaceWithText](Publisher.FindReplace.ReplaceWithText.md)|


---
title: Find.MatchWildcards property (Word)
keywords: vbawd10.chm162529295
f1_keywords:
- vbawd10.chm162529295
ms.prod: word
api_name:
- Word.Find.MatchWildcards
ms.assetid: d2aae410-691e-f718-b888-19e90372d18e
ms.date: 06/08/2017
localization_priority: Normal
---


# Find.MatchWildcards property (Word)

**True** if the text to find contains wildcards. Read/write **Boolean**.


## Syntax

_expression_. `MatchWildcards`

 _expression_ An expression that returns a '[Find](Word.Find.md)' object.


## Remarks

The  **MatchWildcards** property corresponds to the **Use wildcards** check box in the **Find and Replace** dialog box (**Edit** menu).

Use the  **[Text](Word.Find.Text.md)** property of the **Find** object or use the FindText argument with the **[Execute](Word.Find.Execute.md)** method to specify the text to be located in a document.


## Example

This example finds and selects the next three-letter word that begins with "s" and ends with "t."

```vb
With Selection.Find 
 .ClearFormatting 
 .Text = "s*t" 
 .MatchAllWordForms = False 
 .MatchSoundsLike = False 
 .MatchFuzzy = False 
 .MatchWildcards = True 
 .Execute Format:=False, Forward:=True 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
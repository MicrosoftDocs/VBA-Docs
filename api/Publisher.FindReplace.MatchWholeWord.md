---
title: FindReplace.MatchWholeWord property (Publisher)
keywords: vbapb10.chm8323083
f1_keywords:
- vbapb10.chm8323083
ms.prod: publisher
api_name:
- Publisher.FindReplace.MatchWholeWord
ms.assetid: 512d37bc-c900-ee17-8a8e-5875db0a2f85
ms.date: 06/07/2019
localization_priority: Normal
---


# FindReplace.MatchWholeWord property (Publisher)

Sets or returns a **Boolean** that represents whether the whole word will be matched in the search operation. Read/write **Boolean**.


## Syntax

_expression_.**MatchWholeWord**

_expression_ A variable that represents a **[FindReplace](Publisher.FindReplace.md)** object.


## Return value

Boolean


## Remarks

The default value for **MatchWholeWord** is **False**.


## Example

This example selects each occurrence of the word "fact" and applies bold formatting.

```vb
With ActiveDocument.Find 
 .Clear 
 .MatchWholeWord = True 
 .FindText = "fact" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With 

```

<br/>

This example follows the previous example except that whole words will not be matched. Therefore the word "fact" within the word "factory" or "factoid" will also have bold formatting applied.

```vb
With ActiveDocument.Find 
 .Clear 
 .MatchWholeWord = False 
 .FindText = "fact" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
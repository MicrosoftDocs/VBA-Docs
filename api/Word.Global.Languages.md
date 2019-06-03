---
title: Global.Languages property (Word)
keywords: vbawd10.chm163119118
f1_keywords:
- vbawd10.chm163119118
ms.prod: word
api_name:
- Word.Global.Languages
ms.assetid: 6f0d87f8-f0f8-5865-3ba5-2a383c212998
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.Languages property (Word)

Returns a  **Languages** collection that represents the proofing languages listed in the **Language** dialog box.


## Syntax

_expression_. `Languages`

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example returns the full path and file name of the active spelling dictionary.


```vb
Dim dicSpell As Dictionary 
 
Set dicSpell = _ 
 Languages(Selection.LanguageID).ActiveSpellingDictionary 
 
MsgBox dicSpell.Path & Application.PathSeparator & dicSpell.Name
```

This example uses the  _aLang()_ array to store the proofing language names.




```vb
Dim intCount As Integer 
Dim langLoop As Language 
Dim aLang(Languages.Count - 1) As String 
 
intCount = 0 
For Each langLoop In Languages 
 aLang(intCount) = langLoop.NameLocal 
 intCount = intCount + 1 
Next langLoop
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
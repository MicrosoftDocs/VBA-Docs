---
title: Application.Languages property (Word)
keywords: vbawd10.chm158334990
f1_keywords:
- vbawd10.chm158334990
ms.prod: word
api_name:
- Word.Application.Languages
ms.assetid: f81cfcb6-33e2-bb8e-2ac4-b4f9df833946
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Languages property (Word)

Returns a  **[Languages](Word.languages.md)** collection that represents the proofing languages listed in the **Language** dialog box.


## Syntax

_expression_. `Languages`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


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

This example uses the  `aLang()` array to store the proofing language names.




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


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Language.ActiveSpellingDictionary property (Word)
keywords: vbawd10.chm158138383
f1_keywords:
- vbawd10.chm158138383
ms.prod: word
api_name:
- Word.Language.ActiveSpellingDictionary
ms.assetid: a549c07d-e40f-2731-40a0-4d43211cf557
ms.date: 06/08/2017
localization_priority: Normal
---


# Language.ActiveSpellingDictionary property (Word)

Returns a  **[Dictionary](Word.Dictionary.md)** object that represents the active spelling dictionary for the specified language.


## Syntax

_expression_. `ActiveSpellingDictionary`

 _expression_ An expression that returns a '[Language](Word.Language.md)' object.


## Remarks

If there is no spelling dictionary installed for the specified language, this property returns  **Nothing**.


## Example

This example returns the full path and file name of the active spelling dictionary.


```vb
Dim lngLanguage As Long 
Dim dicSpelling As Dictionary 
 
lngLanguage = Selection.LanguageID 
Set dicSpelling = Languages(lngLanguage).ActiveSpellingDictionary 
If dicSpelling Is Nothing Then 
 MsgBox "No spelling dictionary installed!" 
Else 
 MsgBox dicSpelling.Path & Application.PathSeparator _ 
 & dicSpelling.Name 
End If 

```


## See also


[Language Object](Word.Language.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Language.ActiveGrammarDictionary property (Word)
keywords: vbawd10.chm158138381
f1_keywords:
- vbawd10.chm158138381
ms.prod: word
api_name:
- Word.Language.ActiveGrammarDictionary
ms.assetid: 6cded20a-78e3-f01b-9ea8-42134ca5d7c7
ms.date: 06/08/2017
localization_priority: Normal
---


# Language.ActiveGrammarDictionary property (Word)

Returns a  **[Dictionary](Word.Dictionary.md)** object that represents the active grammar dictionary for the specified language. Read-only.


## Syntax

 _expression_. `ActiveGrammarDictionary`

 _expression_ A variable that represents a '[Language](Word.Language.md)' object.


## Remarks

If there is no grammar dictionary installed for the specified language, this property returns  **Nothing**.


## Example

This example displays the full path and file name of the active grammar dictionary.


```vb
Dim lngLanguage As Long 
Dim dicGrammar As Dictionary 
 
lngLanguage = Selection.LanguageID 
Set dicGrammar = Languages(lngLanguage).ActiveGrammarDictionary 
MsgBox dicGrammar.Path & Application.PathSeparator & dicGrammar.Name
```


## See also


[Language Object](Word.Language.md)


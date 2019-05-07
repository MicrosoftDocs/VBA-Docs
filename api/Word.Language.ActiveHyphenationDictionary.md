---
title: Language.ActiveHyphenationDictionary property (Word)
keywords: vbawd10.chm158138382
f1_keywords:
- vbawd10.chm158138382
ms.prod: word
api_name:
- Word.Language.ActiveHyphenationDictionary
ms.assetid: 355462bc-c39e-2e2c-0d2e-af5d4ee8c5a7
ms.date: 06/08/2017
localization_priority: Normal
---


# Language.ActiveHyphenationDictionary property (Word)

Returns a  **[Dictionary](Word.Dictionary.md)** object that represents the active hyphenation dictionary for the specified language. Read-only.


## Syntax

_expression_. `ActiveHyphenationDictionary`

_expression_ A variable that represents a '[Language](Word.Language.md)' object.


## Remarks

If there is no hyphenation dictionary installed for the specified language, this property returns  **Nothing**.


## Example

This example displays the full path and file name of the active hyphenation dictionary.


```vb
Dim lngLanguage As Long 
Dim dicHyphen As Dictionary 
 
lngLanguage = Selection.LanguageID 
Set dicHyphen = Languages(lngLanguage).ActiveHyphenationDictionary 
If dicHyphen Is Nothing Then 
 MsgBox "No hyphenation dictionary installed!" 
Else 
 MsgBox dicHyphen.Path & Application.PathSeparator & dicHyphen.Name 
End If
```


## See also


[Language Object](Word.Language.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
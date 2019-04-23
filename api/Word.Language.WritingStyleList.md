---
title: Language.WritingStyleList property (Word)
keywords: vbawd10.chm158138386
f1_keywords:
- vbawd10.chm158138386
ms.prod: word
api_name:
- Word.Language.WritingStyleList
ms.assetid: 5a91ecaa-dce0-d9ab-0e25-ec9620fa7119
ms.date: 06/08/2017
localization_priority: Normal
---


# Language.WritingStyleList property (Word)

Returns a string array that contains the names of all writing styles available for the specified language. Read-only  **Variant**.


## Syntax

_expression_. `WritingStyleList`

 _expression_ An expression that returns a '[Language](Word.Language.md)' object.


## Example

This example displays each writing style available for U.S. English. Each writing style and its number in the array are also displayed in the Immediate window of the Visual Basic editor.


```vb
Sub WritingStyles() 
 Dim WrStyles As Variant 
 Dim i As Integer 
 
 WrStyles = Languages(wdEnglishUS).WritingStyleList 
 For i = 1 To UBound(WrStyles) 
 MsgBox WrStyles(i) 
 Debug.Print WrStyles(i) & " [" & Trim(Str$(i)) & "]" 
 Next i 
End Sub
```


## See also


[Language Object](Word.Language.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
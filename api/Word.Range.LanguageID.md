---
title: Range.LanguageID property (Word)
keywords: vbawd10.chm157155481
f1_keywords:
- vbawd10.chm157155481
ms.prod: word
api_name:
- Word.Range.LanguageID
ms.assetid: dc163c7b-8a44-4b8a-5674-845984f1b682
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.LanguageID property (Word)

Returns or sets a  **[WdLanguageID](Word.WdLanguageID.md)** constant that represents the language for the specified range. Read/write.


## Syntax

_expression_. `LanguageID`

 _expression_ An expression that represents a **[Range](Word.Range.md)** object.


## Remarks

Some of the **WdLanguageID** constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example formats the second paragraph in the active document as French and then adds a new custom dictionary that will be used on the French text.


```vb
ActiveDocument.Paragraphs(2).Range.LanguageID = wdFrench 
Set myDictionary = CustomDictionaries.Add(Filename:="French.dic") 
With myDictionary 
 .LanguageSpecific = True 
 .LanguageID = wdFrench 
End With
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
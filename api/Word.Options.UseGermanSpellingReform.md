---
title: Options.UseGermanSpellingReform property (Word)
keywords: vbawd10.chm162988447
f1_keywords:
- vbawd10.chm162988447
ms.prod: word
api_name:
- Word.Options.UseGermanSpellingReform
ms.assetid: 5ab20040-7247-f402-c246-e13c1ba0cb30
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.UseGermanSpellingReform property (Word)

 **True** if Microsoft Word uses the German post-reform spelling rules when checking spelling. Read/write **Boolean**.


## Syntax

 _expression_. `UseGermanSpellingReform`

 _expression_ An expression that returns an '[Options](Word.Options.md)' object.


## Remarks

This property may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example sets Word to use the post-reform rules for checking spelling in German.


```vb
Options.UseGermanSpellingReform = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
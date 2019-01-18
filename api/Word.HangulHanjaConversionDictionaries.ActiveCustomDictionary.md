---
title: HangulHanjaConversionDictionaries.ActiveCustomDictionary property (Word)
keywords: vbawd10.chm165675011
f1_keywords:
- vbawd10.chm165675011
ms.prod: word
api_name:
- Word.HangulHanjaConversionDictionaries.ActiveCustomDictionary
ms.assetid: 3e1d8fd9-eee8-eb18-f4db-6a9e5379436e
ms.date: 06/08/2017
localization_priority: Normal
---


# HangulHanjaConversionDictionaries.ActiveCustomDictionary property (Word)

Returns or sets a  **[Dictionary](Word.Dictionary.md)** object that represents the custom dictionary to which words will be added. Read/write.


## Syntax

 _expression_. `ActiveCustomDictionary`

 _expression_ A variable that represents a '[HangulHanjaConversionDictionaries](Word.hangulhanjaconversiondictionaries.md)' collection.


## Example

This example displays the full path and file name of the active custom dictionary.


```vb
Set dicCustom = Application.CustomDictionaries.ActiveCustomDictionary 
MsgBox dicCustom.Path & Application.PathSeparator & dicCustom.Name
```

This example clears all existing custom dictionaries, adds a custom dictionary named "Home.dic," and then loads the new dictionary.




```vb
Dim dicCustom As Dictionary 
 
Application.CustomDictionaries.ClearAll 
 
Set dicCustom = Application.CustomDictionaries _ 
 .Add(FileName:="C:\Program Files" _ 
 & "\Microsoft Office\Office\Home.dic") 
Application.CustomDictionaries.ActiveCustomDictionary = dicCustom
```


## See also


[HangulHanjaConversionDictionaries Collection Object](Word.hangulhanjaconversiondictionaries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
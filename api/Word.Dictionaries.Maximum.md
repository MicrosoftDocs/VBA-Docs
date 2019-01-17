---
title: Dictionaries.Maximum property (Word)
keywords: vbawd10.chm162267138
f1_keywords:
- vbawd10.chm162267138
ms.prod: word
api_name:
- Word.Dictionaries.Maximum
ms.assetid: fa9f31e0-1965-5d96-568b-e0b8049127e3
ms.date: 06/08/2017
localization_priority: Normal
---


# Dictionaries.Maximum property (Word)

Returns the maximum number of custom or conversion dictionaries allowed. Read-only  **Long**.


## Syntax

 _expression_. `Maximum`

 _expression_ Required. A variable that represents a '[Dictionaries](Word.dictionaries.md)' collection.


## Example

This example displays a message if the number of custom dictionaries is equal to the maximum number allowed. If the maximum number has not been reached, a custom dictionary named "MyDictionary.dic" is added.


```vb
If CustomDictionaries.Count = CustomDictionaries.Maximum Then 
 MsgBox "Cannot add another dictionary file" 
Else 
 CustomDictionaries.Add "MyDictionary.dic" 
End If
```


## See also


[Dictionaries Collection Object](Word.dictionaries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
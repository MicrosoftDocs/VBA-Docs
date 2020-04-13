---
title: Application.CustomDictionaries property (Word)
keywords: vbawd10.chm158335071
f1_keywords:
- vbawd10.chm158335071
ms.prod: word
api_name:
- Word.Application.CustomDictionaries
ms.assetid: 1c6dca90-70f0-6b52-72d1-debda33d2ba0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CustomDictionaries property (Word)

Returns a  **[Dictionaries](Word.dictionaries.md)** object that represents the collection of active custom dictionaries. Read-only.


## Syntax

_expression_. `CustomDictionaries`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

Active custom dictionaries are marked with a check in the **Custom Dictionaries** dialog box. For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example adds a new, blank custom dictionary to the collection. The path and file name of the new custom dictionary are then displayed in a message box.


```vb
Dim dicHome As Dictionary 
Set dicHome = CustomDictionaries.Add(Filename:="Home.dic") 
Msgbox dicHome.Path & Application.PathSeparator & dicHome.Name
```

This example removes all custom dictionaries so that no custom dictionaries are active. The custom dictionary files aren't deleted, though.




```vb
CustomDictionaries.ClearAll
```

This example displays the name of each custom dictionary in the collection.




```vb
Dim dicLoop As Dictionary 
 
For Each dicLoop In CustomDictionaries 
 MsgBox dicLoop.Name 
Next dicLoop
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
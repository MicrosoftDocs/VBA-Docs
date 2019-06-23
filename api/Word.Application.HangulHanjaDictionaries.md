---
title: Application.HangulHanjaDictionaries property (Word)
keywords: vbawd10.chm158335086
f1_keywords:
- vbawd10.chm158335086
ms.prod: word
api_name:
- Word.Application.HangulHanjaDictionaries
ms.assetid: 453e2a77-f363-5afc-d9a3-26f8b6516b4c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.HangulHanjaDictionaries property (Word)

Returns a **[HangulHanjaConversionDictionaries](Word.hangulhanjaconversiondictionaries.md)** collection that represents all the active custom conversion dictionaries.


## Syntax

_expression_. `HangulHanjaDictionaries`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

Active custom conversion dictionaries are marked with a check in the **Custom Dictionaries** dialog box. Click **Options**, click the **Spelling & Grammar** tab, and then click the **Custom Dictionaries** button.

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example adds a new, blank custom dictionary to the collection. The path and file name of the new custom dictionary are then displayed in a message box.


```vb
Set myHome = _ 
 HangulHanjaDictionaries.Add(Filename:="Home.hhd") 
Msgbox myHome.Path & Application.PathSeparator _ 
 & myHome.Name
```

This example deactivates all custom dictionaries but does not delete the custom dictionary files.




```vb
HangulHanjaDictionaries.ClearAll
```

This example displays the name of each custom dictionary in the collection.




```vb
For Each di In HangulHanjaDictionaries 
 MsgBox di.Name 
Next di
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
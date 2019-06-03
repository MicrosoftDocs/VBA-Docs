---
title: Global.AddIns property (Word)
keywords: vbawd10.chm163119126
f1_keywords:
- vbawd10.chm163119126
ms.prod: word
api_name:
- Word.Global.AddIns
ms.assetid: 21b0d291-aa8c-28c0-ef5e-6a566d17da9d
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.AddIns property (Word)

Returns an  **[AddIns](Word.addins.md)** collection that represents all available add-ins, regardless of whether they're currently loaded. Read-only.


## Syntax

_expression_. `AddIns`

_expression_ A variable that represents a '[Global](Word.Global.md)' object.


## Remarks

The  **[AddIns](Word.addins.md)** collection includes the global templates and Word add-in libraries (WLLs) listed in the **Templates and Add-ins** dialog box (**Tools** menu).

 For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example returns the total number of add-ins.


```vb
Dim intAddIns as Integer 
 
intAddIns = AddIns.Count
```

This example displays the name of each add-in in the Addins collection.




```vb
Dim addinLoop as AddIn 
 
For Each addinLoop In AddIns 
 MsgBox addinLoop.Name 
Next addinLoop
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
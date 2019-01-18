---
title: RemoveAll method (Visual Basic for Applications)
keywords: vblr6.chm2181953
f1_keywords:
- vblr6.chm2181953
ms.prod: office
api_name:
- Office.RemoveAll
ms.assetid: 70edc5db-1f44-cfa5-cf22-13a9ce33a954
ms.date: 12/14/2018
localization_priority: Normal
---


# RemoveAll method

The **RemoveAll** method removes all key, item pairs from a **[Dictionary](dictionary-object.md)** object.

## Syntax

_object_.**RemoveAll**

The _object_ is always the name of a **Dictionary** object.

## Remarks

The following code illustrates use of the **RemoveAll** method.


```vb
Dim a, d, i             'Create some variables
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
...
a = d.RemoveAll         'Clear the dictionary
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
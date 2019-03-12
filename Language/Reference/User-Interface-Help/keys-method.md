---
title: Keys method (Visual Basic for Applications)
keywords: vblr6.chm2181951
f1_keywords:
- vblr6.chm2181951
ms.prod: office
api_name:
- Office.Keys
ms.assetid: d5ec76fc-d293-264b-7b84-1d9e7cac170c
ms.date: 12/14/2018
localization_priority: Normal
---


# Keys method

Returns an array containing all existing keys in a **[Dictionary](dictionary-object.md)** object.

## Syntax

_object_.**Keys**

The _object_ is always the name of a **Dictionary** object.

## Remarks

The following code illustrates use of the **Keys** method:

```vb
Dim a, d, i             'Create some variables
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items.
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
a = d.keys              'Get the keys
For i = 0 To d.Count -1 'Iterate the array
    Print a(i)          'Print key
Next
...

```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

---
title: Items method (Visual Basic for Applications)
keywords: vblr6.chm2181950
f1_keywords:
- vblr6.chm2181950
ms.prod: office
api_name:
- Office.Items
ms.assetid: ba058f8d-d0b1-c93f-95fc-7d2e8744808c
ms.date: 12/14/2018
localization_priority: Normal
---


# Items method

Returns an array containing all the items in a **[Dictionary](dictionary-object.md)** object.

## Syntax

_object_.**Items**

The _object_ is always the name of a **Dictionary** object.

## Remarks

The following code illustrates use of the **Items** method:

```vb
Dim a, d, i             'Create some variables
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
a = d.Items             'Get the items
For i = 0 To d.Count -1 'Iterate the array
    Print a(i)          'Print item
Next
...

```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
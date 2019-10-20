---
title: Count property (Dictionary object)
keywords: vblr6.chm2181945
f1_keywords:
- vblr6.chm2181945
ms.prod: office
ms.assetid: b64c41d8-3fe3-3a69-0949-a1d1956be12f
ms.date: 04/18/2019
localization_priority: Normal
---


# Count property

Returns a **Long** (long integer) containing the number of items in a collection or **[Dictionary](dictionary-object.md)** object. Read-only.

## Syntax

_object_.**Count**

The _object_ is always the name of one of the items in the **Applies To** list.

## Remarks

The following code illustrates use of the **Count** property.

```vb
Dim a, d, i             'Create some variables
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items.
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
a = d.Keys              'Get the keys
For i = 0 To d.Count -1 'Iterate the array
    Print a(i)          'Print key
Next
...

```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

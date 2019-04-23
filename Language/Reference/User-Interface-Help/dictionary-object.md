---
title: Dictionary object
keywords: vblr6.chm2181922
f1_keywords:
- vblr6.chm2181922
ms.prod: office
api_name:
- Office.Dictionary
ms.assetid: 718dbcd4-63bc-3a75-fa55-7d1e8c65e8b9
ms.date: 04/02/2019
localization_priority: Normal
---


# Dictionary object

Object that stores data key/item pairs.

## Syntax

**Scripting.Dictionary**

## Remarks

A **Dictionary** object is the equivalent of a PERL associative array. Items, which can be any form of data, are stored in the array. Each item is associated with a unique key. The key is used to retrieve an individual item and is usually an integer or a string, but can be anything except an array.

The following code illustrates how to create a **Dictionary** object.

```vb
Dim d                   'Create a variable
Set d = CreateObject("Scripting.Dictionary")
d.Add "a", "Athens"     'Add some keys and items
d.Add "b", "Belgrade"
d.Add "c", "Cairo"
...
```

## Methods

|Method|Description|
|:-----|:----------|
|[Add](add-method-dictionary.md)|Adds a new key/item pair to a **Dictionary** object. |
|[Exists](exists-method.md)|Returns a Boolean value that indicates whether a specified key exists in the **Dictionary** object. |
|[Items](items-method.md)|Returns an array of all the items in a **Dictionary** object. |
|[Keys](keys-method.md)|Returns an array of all the keys in a **Dictionary** object. |
|[Remove](remove-method-dictionary-object.md)|Removes one specified key/item pair from the **Dictionary** object. |
|[RemoveAll](removeall-method.md)|Removes all the key/item pairs in the **Dictionary** object. |


## Properties

|Property|Description|
|:-------|:----------|
|[CompareMode](comparemode-property.md)|Sets or returns the comparison mode for comparing keys in a **Dictionary** object. |
|[Count](count-property-dictionary-object.md)|Returns the number of key/item pairs in a **Dictionary** object. |
|[Item](item-property-dictionary-object.md)|Sets or returns the value of an item in a **Dictionary** object. |
|[Key](key-property.md)|Sets a new key value for an existing key value in a **Dictionary** object. |

## See also

- [Dictionary object (Windows Scripting previous version)](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/windows-scripting/x4k5wbx4(v%3dvs.84))
- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Object library reference for Office (members, properties, methods)](../../../api/overview/library-reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

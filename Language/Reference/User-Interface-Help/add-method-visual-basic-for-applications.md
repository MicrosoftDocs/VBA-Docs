---
title: Add method (Visual Basic for Applications)
keywords: vblr6.chm1014017
f1_keywords:
- vblr6.chm1014017
ms.assetid: c9e9eb2e-42b1-9821-67ab-2f68fb87d1d0
ms.date: 12/14/2018
ms.localizationpriority: medium
---


# Add method (VBA)

Adds a [member](../../Glossary/vbe-glossary.md#member) to a **[Collection](collection-object.md)** object.

## Syntax

_object_.**Add** _item_, _key_, _before_, _after_

The **Add** method syntax has the following object qualifier and [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
|_item_|Required. An [expression](../../Glossary/vbe-glossary.md#expression) of any type that specifies the member to add to the [collection](../../Glossary/vbe-glossary.md#collection).|
|_key_|Optional. A unique [string expression](../../Glossary/vbe-glossary.md#string-expression) that specifies a key string that can be used, instead of a positional index, to access a member of the collection.|
|_before_|Optional. An expression that specifies a relative position in the collection. The member to be added is placed in the collection before the member identified by the _before_ [argument](../../Glossary/vbe-glossary.md#argument). If a [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression), _before_ must be a number from 1 to the value of the collection's **[Count](count-property-visual-basic-for-applications.md)** property. If a string expression, _before_ must correspond to the _key_ specified when the member being referred to was added to the collection. You can specify a _before_ position or an _after_ position, but not both.|
|_after_|Optional. An expression that specifies a relative position in the collection. The member to be added is placed in the collection after the member identified by the _after_ argument. If numeric, _after_ must be a number from 1 to the value of the collection's **Count** property. If a string, _after_ must correspond to the _key_ specified when the member referred to was added to the collection. You can specify a _before_ position or an _after_ position, but not both.|

## Remarks

Whether the _before_ or _after_ argument is a string expression or numeric expression, it must refer to an existing member of the collection, or an error occurs.

An error also occurs if a specified _key_ duplicates the _key_ for an existing member of the collection.

## Example

This example uses the **Add** method to add strings to a collection both with and without a _key_. The **[Item](item-method-visual-basic-for-applications.md)** method is used implicitly to retrieve each string.

```vb
Dim c As Collection
Set c = New Collection

c.Add "a"
c.Add "c", "CC"
c.Add "b", "BB", 2
c.Add "d"

Debug.Print c(1) ' --> prints "a"
Debug.Print c(2) ' --> prints "b"
Debug.Print c(3) ' --> prints "c"

Debug.Print c("BB") ' --> prints "b"
Debug.Print c("d") ' --> error (no key was specified for this element - a positional index must be used)
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

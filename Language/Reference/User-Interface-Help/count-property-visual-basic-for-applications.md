---
title: Count property (Visual Basic for Applications)
keywords: vblr6.chm1014018
f1_keywords:
- vblr6.chm1014018
ms.prod: office
ms.assetid: 319907c0-9c68-0e24-cc76-c16d4386269e
ms.date: 12/19/2018
localization_priority: Normal
---


# Count property (VBA)

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) (long integer) containing the number of objects in a [collection](../../Glossary/vbe-glossary.md#collection). Read-only.

## Example

This example uses the **[Collection](collection-object.md)** object's **Count** property to specify how many iterations are required to remove all the elements of the collection called `MyClasses`. When collections are numerically indexed, the base is 1 by default. Because collections are reindexed automatically when a removal is made, the following code removes the first member on each iteration.

```vb
Dim Num, MyClasses
For Num = 1 To MyClasses.Count    ' Remove name from the collection.
    MyClasses.Remove 1    ' Default collection numeric indexes
Next    ' begin at 1.
```


## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

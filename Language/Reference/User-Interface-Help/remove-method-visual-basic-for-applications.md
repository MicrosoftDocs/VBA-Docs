---
title: Remove Method (Visual Basic for Applications)
keywords: vblr6.chm1014020
f1_keywords:
- vblr6.chm1014020
ms.prod: office
ms.assetid: ad45eba6-eb95-3cdc-03c2-7c94e8a38d48
ms.date: 06/08/2017
---


# Remove Method (Visual Basic for Applications)



Removes a [member](../../Glossary/vbe-glossary.md#member) from a **Collection** object.

## Syntax

_object_**.Remove**_index_
The  **Remove** method syntax has the following object qualifier and part:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the Applies To list.|
| _index_|Required. An [expression](../../Glossary/vbe-glossary.md#expression) that specifies the position of a member of the[collection](../../Glossary/vbe-glossary.md#collection). If a [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression),  _index_ must be a number from 1 to the value of the collection's **Count**[property](../../Glossary/vbe-glossary.md#property). If a [string expression](../../Glossary/vbe-glossary.md#string-expression),  _index_ must correspond to the **_key_**[argument](../../Glossary/vbe-glossary.md#argument) specified when the member referred to was added to the collection.|

## Remarks

<<<<<<< HEAD
=======
## Remarks

>>>>>>> 54e0a75f224118db0d26fc9363ad519ad35ec788
If the value provided as  _index_ doesn't match an existing member of the collection, an error occurs.

## Example

This example illustrates the use of the  **Remove** method to remove objects from a **Collection** object, `MyClasses`. This code removes the object whose index is 1 on each iteration of the loop.


```vb
Dim Num, MyClasses
For Num = 1 To MyClasses.Count    
    MyClasses.Remove 1    ' Remove the first object each time
            ' through the loop until there are 
            ' no objects left in the collection.
Next Num

```



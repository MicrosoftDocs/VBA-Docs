---
title: DropDown.ListEntries property (Word)
keywords: vbawd10.chm153419779
f1_keywords:
- vbawd10.chm153419779
ms.prod: word
api_name:
- Word.DropDown.ListEntries
ms.assetid: 87235132-0ff6-e8d7-1efc-1df4a9816b2f
ms.date: 06/08/2017
localization_priority: Normal
---


# DropDown.ListEntries property (Word)

Returns a  **[ListEntries](Word.listentries.md)** collection that represents all the items in a **DropDown** object.


## Syntax

_expression_. `ListEntries`

 _expression_ An expression that returns a '[DropDown](Word.DropDown.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example retrieves the text of the active item from the drop-down form field named "DropDown1."


```vb
Set myField = ActiveDocument.FormFields("DropDown1").DropDown 
num = myField.Value 
myName = myField.ListEntries(num).Name
```

This example retrieves the total number of items in the active drop-down form field (the document should be protected for forms). If there are two or more items, this example sets the second item as the active item.




```vb
Set myField = Selection.FormFields(1) 
If myfield.Type = wdFieldFormDropDown Then 
 num = myField.DropDown.ListEntries.Count 
 If num >= 2 Then myField.DropDown.Value = 2 
End If
```


## See also


[DropDown Object](Word.DropDown.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
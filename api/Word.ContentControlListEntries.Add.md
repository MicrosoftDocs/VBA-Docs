---
title: ContentControlListEntries.Add method (Word)
keywords: vbawd10.chm230948970
f1_keywords:
- vbawd10.chm230948970
ms.prod: word
api_name:
- Word.ContentControlListEntries.Add
ms.assetid: 159747c0-279c-f0ee-62d9-f2f01865c083
ms.date: 06/08/2017
localization_priority: Normal
---


# ContentControlListEntries.Add method (Word)

Adds a new list item to a drop-down list or combo box content control and returns a  **[ContentControlListEntry](Word.ContentControlListEntry.md)** object.


## Syntax

_expression_.**Add** (_Text_, _Value_, _Index_)

 _expression_ An expression that returns a [ContentControlListEntries](./Word.ContentControlListEntries.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|Specifies the display text for the list item. Corresponds to the  **[Text](Word.ContentControlListEntry.Text.md)** property for a **ContentControlListEntry** object.|
| _Value_|Optional| **String**|Specifies the value of the list item. Corresponds to the  **[Value](Word.ContentControlListEntry.Value.md)** property for a **ContentControlListEntry** object. If omitted, the **Value** property is equal to the **Text** property.|
| _Index_|Optional| **Long**|Specifies the ordinal position of the new item in the list. If an item exists at the position specified, the existing item is pushed down in the list. If omitted, the new item is added to the end of the list.|

## Return value

ContentControlListEntry


## Remarks

List entries must have unique display names. Attempting to add a list item that already exists raises a run-time error.


## Example

The following code example creates a new drop-down list content control and then adds several items to it.


```vb
Dim objCC As ContentControl 
Dim objLE As ContentControlListEntry 
Dim objMap As XMLMapping 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList) 
 
'List items 
objCC.DropdownListEntries.Add "Cat" 
objCC.DropdownListEntries.Add "Dog" 
objCC.DropdownListEntries.Add "Equine" 
objCC.DropdownListEntries.Add "Monkey" 
objCC.DropdownListEntries.Add "Snake" 
objCC.DropdownListEntries.Add "Other"
```


## See also


[ContentControlListEntries Collection](Word.ContentControlListEntries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
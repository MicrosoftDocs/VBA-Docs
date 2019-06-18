---
title: WebListBoxItems.AddItem method (Publisher)
keywords: vbapb10.chm4128772
f1_keywords:
- vbapb10.chm4128772
ms.prod: publisher
api_name:
- Publisher.WebListBoxItems.AddItem
ms.assetid: 1c3af4d1-ed0b-60c6-b607-17712612cec2
ms.date: 06/18/2019
localization_priority: Normal
---


# WebListBoxItems.AddItem method (Publisher)

Adds list items to a web list box control.


## Syntax

_expression_.**AddItem** (_Item_, _Index_, _SelectState_, _ItemValue_)

_expression_ A variable that represents a **[WebListBoxItems](Publisher.WebListBoxItems.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Item_|Required| **String**|The name of the item as it appears in the list.|
|_Index_|Optional| **Long**|The number of the list item. If _Index_ is not specified or if it is out of range of the indices of existing list box items, the new item is added to the end of the list box. Otherwise, the new item is inserted at the position specified by _Index_, and the index position of all the items after it are increased by one.|
|_SelectState_|Optional| **Boolean**| **True** if the item is selected when the list box is initially displayed. The default value is **False**.|
|_ItemValue_|Optional| **String**|The value of the list box item. If not specified, the new item's value is the same as the item name.|

## Remarks

When you programmatically create a new web list box, it contains three items. Use the **[Delete](Publisher.WebListBoxItems.Delete.md)** method to remove them from the list.


## Example

This example creates a new list box control in the active publication, removes the three default list items, and then adds several items to it.

```vb
Sub AddListBoxItems() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlListBox, Left:=100, _ 
 Top:=100, Width:=150, Height:=100) 
 With .WebListBox.ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next 
 .AddItem Item:="Green" 
 .AddItem Item:="Yellow" 
 .AddItem Item:="Red" 
 .AddItem Item:="Blue" 
 .AddItem Item:="Purple" 
 .AddItem Item:="Chartreuse" 
 .AddItem Item:="Pink" 
 .AddItem Item:="Olive" 
 End With 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
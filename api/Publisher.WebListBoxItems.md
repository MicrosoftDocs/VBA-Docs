---
title: WebListBoxItems object (Publisher)
keywords: vbapb10.chm4194303
f1_keywords:
- vbapb10.chm4194303
ms.prod: publisher
api_name:
- Publisher.WebListBoxItems
ms.assetid: 6d1b6755-426b-b518-c95c-7b30f9acceba
ms.date: 06/04/2019
localization_priority: Normal
---


# WebListBoxItems object (Publisher)

Represents the items in a web list box control.
 
## Remarks

Use the **[ListBoxItems](Publisher.WebListBox.ListBoxItems.md)** property of the **WebListBox** object to access the items in a web list box. 

Use the **AddItem** method to add items to a web list box. 

## Example

This example creates a new web list box and adds several items to it. Note that when initially created, a web list box control contains three default items. This example includes a routine that deletes the default list box items before adding new items.

```vb
Sub CreateWebListBox() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlListBox, Left:=100, _ 
 Top:=150, Width:=300, Height:=72).WebListBox 
 .MultiSelect = msoFalse 
 With .ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next 
 .AddItem Item:="Green" 
 .AddItem Item:="Purple" 
 .AddItem Item:="Red" 
 .AddItem Item:="Black" 
 End With 
 End With 
 End With 
End Sub
```

## Methods

- [AddItem](Publisher.WebListBoxItems.AddItem.md)
- [Delete](Publisher.WebListBoxItems.Delete.md)
- [Item](Publisher.WebListBoxItems.Item.md)
- [Selected](Publisher.WebListBoxItems.Selected.md)

## Properties

- [Application](Publisher.WebListBoxItems.Application.md)
- [Count](Publisher.WebListBoxItems.Count.md)
- [Parent](Publisher.WebListBoxItems.Parent.md)


## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
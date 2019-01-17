---
title: WebListBox Object (Publisher)
keywords: vbapb10.chm4128767
f1_keywords:
- vbapb10.chm4128767
ms.prod: publisher
api_name:
- Publisher.WebListBox
ms.assetid: 0ba881f8-95cf-c536-7fa8-05714348577d
ms.date: 06/08/2017
localization_priority: Normal
---


# WebListBox Object (Publisher)

Represents a Web list box control. The  **WebListBox** object is a member of the **Shape** object.
 


## Example

Use the  **[AddWebControl](Publisher.Shapes.AddWebControl.md)** method to create a new Web list box. Use the **[WebListBox](Publisher.Shape.WebListBox.md)** property to access a Web list box control shape. Use the **[AddItem](Publisher.WebListBoxItems.AddItem.md)** method of the **[WebListBoxItems](Publisher.WebListBoxItems.md)** object to add items to a Web list box. This example creates a new Web list box and adds several items to it. Note that when initially created, a Web list box control contains three default items. This example includes a routine that deletes the default list box items before adding new items.
 

 

 

 

 **Note**  When you create a Web list box, its initial width is 300 points. However, Microsoft Publisher automatically changes this width based on the width of the items in the list.
 




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


## Properties



|Name|
|:-----|
|[Application](Publisher.WebListBox.Application.md)|
|[ListBoxItems](Publisher.WebListBox.ListBoxItems.md)|
|[MultiSelect](Publisher.WebListBox.MultiSelect.md)|
|[Parent](Publisher.WebListBox.Parent.md)|
|[ReturnDataLabel](Publisher.WebListBox.ReturnDataLabel.md)|


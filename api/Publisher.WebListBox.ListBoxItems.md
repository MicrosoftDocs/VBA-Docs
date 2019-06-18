---
title: WebListBox.ListBoxItems property (Publisher)
keywords: vbapb10.chm4063235
f1_keywords:
- vbapb10.chm4063235
ms.prod: publisher
api_name:
- Publisher.WebListBox.ListBoxItems
ms.assetid: 642a4592-35af-99fa-ee96-6bd8517c618f
ms.date: 06/18/2019
localization_priority: Normal
---


# WebListBox.ListBoxItems property (Publisher)

Returns a **[WebListBoxItems](Publisher.WebListBoxItems.md)** object that represents the items in a web list box control.


## Syntax

_expression_.**ListBoxItems**

_expression_ A variable that represents a **[WebListBox](Publisher.WebListBox.md)** object.


## Return value

WebListBoxItems


## Example

This example creates a new web list box control and adds five new list items to it.

```vb
Sub NewListBoxItems() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlListBox, Left:=100, _ 
 Top:=100, Width:=150, Height:=100).WebListBox 
 .MultiSelect = msoTrue 
 With .ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next 
 .AddItem Item:="Yellow" 
 .AddItem Item:="Red" 
 .AddItem Item:="Blue" 
 .AddItem Item:="Green" 
 .AddItem Item:="Black" 
 End With 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
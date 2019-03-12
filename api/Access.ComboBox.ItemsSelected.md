---
title: ComboBox.ItemsSelected property (Access)
keywords: vbaac10.chm11495
f1_keywords:
- vbaac10.chm11495
ms.prod: access
api_name:
- Access.ComboBox.ItemsSelected
ms.assetid: 7e4f6f12-3d97-b36a-1211-8c95b43642e6
ms.date: 03/01/2019
localization_priority: Normal
---


# ComboBox.ItemsSelected property (Access)

You can use the **ItemsSelected** property to return a read-only reference to the hidden **ItemsSelected** collection. This hidden collection can be used to access data in the selected rows of a multiselect combo box control.


## Syntax

_expression_.**ItemsSelected**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

The **ItemsSelected** collection is unlike other collections in that it is a collection of **Variants** rather than of objects. Each **Variant** is an integer index referring to a selected row in a list box or combo box.

Use the **ItemsSelected** collection in conjunction with the **Column** property or the **ItemData** property to retrieve data from selected rows in a list box or combo box. You can list the **ItemsSelected** collection by using the **For Each...Next** statement.

For example, if you have an **Employees** list box on a form, you can list the **ItemsSelected** collection and use the control's **ItemData** property to return the value of the bound column for each selected row in the list box.

The **ItemsSelected** collection has two properties, the **Count** and **Item** properties, and no methods.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

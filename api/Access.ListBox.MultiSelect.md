---
title: ListBox.MultiSelect property (Access)
keywords: vbaac10.chm11237,vbaac10.chm4432
f1_keywords:
- vbaac10.chm11237,vbaac10.chm4432
ms.prod: access
api_name:
- Access.ListBox.MultiSelect
ms.assetid: 7115a913-1696-03b4-c88b-0626da1d587a
ms.date: 03/22/2019
localization_priority: Normal
---


# ListBox.MultiSelect property (Access)

You can use the **MultiSelect** property to specify whether a user can make multiple selections in a list box on a form and how the multiple selections can be made. Read/write **Byte**.


## Syntax

_expression_.**MultiSelect**

_expression_ A variable that represents a **[ListBox](Access.ListBox.md)** object.


## Remarks

The **MultiSelect** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|None|0|(Default) Multiple selection isn't allowed.|
|Simple|1|Multiple items are selected or deselected by choosing them with the mouse or by pressing the Spacebar.|
|Extended|2|Multiple items are selected by holding down Shift and choosing them with the mouse, or by holding down Shift and pressing an arrow key to extend the selection from the previously selected item to the current item. You can also select items by dragging with the mouse. Holding down Ctrl and choosing an item selects or deselects that item.|

This property can be set only in form Design view.

You can use the **ListIndex** property to return the index number for the selected item. When the **MultiSelect** property is set to Extended or Simple, you can use the list box's **Selected** property or **ItemsSelected** collection to determine the items that are selected. In addition, when the **MultiSelect** property is set to Extended or Simple, the value of the list box control will always be **null**.

If the **MultiSelect** property is set to Extended, requerying the list box clears any selections made by the user.


## Example

To return the value of the **MultiSelect** property for a list box named **Country** on the **Order Entry** form, you can use the following.

```vb
Dim b As Byte b = Forms("Order Entry").Controls("Country").MultiSelect
```

<br/>

To set the **MultiSelect** property, you can use the following.

```vb
Forms("Order Entry").Controls("Country").MultiSelect = 2 ' Extended.
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

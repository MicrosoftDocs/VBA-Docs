---
title: Control.Dropdown method (Access)
keywords: vbaac10.chm10135
f1_keywords:
- vbaac10.chm10135
ms.prod: access
api_name:
- Access.Control.Dropdown
ms.assetid: 45957d42-3e81-f7eb-9579-e5e75c833f59
ms.date: 02/28/2019
localization_priority: Normal
---


# Control.Dropdown method (Access)

You can use the **Dropdown** method to force the list in the specified combo box to drop down.


## Syntax

_expression_.**Dropdown**

_expression_ A variable that represents a **[Control](Access.Control.md)** object.


## Return value

Nothing


## Remarks

For example, you can use this method to cause a combo box listing vendor codes to drop down when the vendor code control receives the focus during data entry.

If the specified combo box control doesn't have the focus, an error occurs. The use of this method is identical to pressing the F4 key when the control has the focus.


## Example

The following example shows how you can use the **Dropdown** method within the **GotFocus** event procedure to force a combo box named **SupplierID** to drop down when it receives the focus.

```vb
Private Sub SupplierID_GotFocus() 
 Me!SupplierID.Dropdown 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
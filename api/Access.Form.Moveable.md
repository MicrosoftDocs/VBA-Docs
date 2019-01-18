---
title: Form.Moveable property (Access)
keywords: vbaac10.chm13524
f1_keywords:
- vbaac10.chm13524
ms.prod: access
api_name:
- Access.Form.Moveable
ms.assetid: ad0db2eb-9905-15d9-7a96-e61cefd12842
ms.date: 06/08/2017
localization_priority: Normal
---


# Form.Moveable property (Access)

Returns or sets a  **Boolean** indicating whether the specified form can be moved by the user; **True** if it can be moved. Read/write.


## Syntax

_expression_. `Moveable`

_expression_ A variable that represents a [Form](Access.Form.md) object.


## Remarks

You can use the  **Move** method to programmatically move a form or report regardless of the value of the **Moveable** property.


## Example

The following example determines whether or not the first form in the current project can be moved.


```vb
If Forms(0).Moveable Then 
 MsgBox "You may move the form." 
Else 
 MsgBox "The form cannot be moved." 
End If 

```


## See also


[Form Object](Access.Form.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
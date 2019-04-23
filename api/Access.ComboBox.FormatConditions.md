---
title: ComboBox.FormatConditions property (Access)
keywords: vbaac10.chm11369
f1_keywords:
- vbaac10.chm11369
ms.prod: access
api_name:
- Access.ComboBox.FormatConditions
ms.assetid: 0eeb11b4-453b-4a00-0a1f-92e3108ab2b9
ms.date: 03/01/2019
localization_priority: Normal
---


# ComboBox.FormatConditions property (Access)

You can use the **FormatConditions** property to return a read-only reference to the **[FormatConditions](Access.FormatConditions.md)** collection and its related properties.


## Syntax

_expression_.**FormatConditions**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Example

The following example sets format properties for an existing conditional format for the **Textbox1** control.

```vb
With Forms("forms1").Controls("Textbox1").FormatConditions(1) 
 .BackColor = RGB(255,255,255) 
 .FontBold = True 
 .ForeColor = RGB(255,0,0) 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
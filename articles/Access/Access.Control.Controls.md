---
title: Control.Controls Property (Access)
keywords: vbaac10.chm10150
f1_keywords:
- vbaac10.chm10150
ms.prod: access
api_name:
- Access.Control.Controls
ms.assetid: 81b01d02-c346-8750-cc8a-4623f24219f6
ms.date: 06/08/2017
---


# Control.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **Control** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[Control Object](Access.Control.md)


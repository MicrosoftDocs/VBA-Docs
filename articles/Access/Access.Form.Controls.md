---
title: Form.Controls Property (Access)
keywords: vbaac10.chm13508
f1_keywords:
- vbaac10.chm13508
ms.prod: access
api_name:
- Access.Form.Controls
ms.assetid: 08a31b50-b644-5912-d784-130f58298dd0
ms.date: 06/08/2017
---


# Form.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **Form** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[Form Object](Access.Form.md)


---
title: Image.Controls Property (Access)
keywords: vbaac10.chm10361
f1_keywords:
- vbaac10.chm10361
ms.prod: access
api_name:
- Access.Image.Controls
ms.assetid: b6313b26-4254-fafb-923b-ef9d2b9fc0f5
ms.date: 06/08/2017
---


# Image.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents an **Image** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[Image Object](Access.Image.md)


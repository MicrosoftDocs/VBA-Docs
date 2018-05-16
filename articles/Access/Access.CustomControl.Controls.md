---
title: CustomControl.Controls Property (Access)
keywords: vbaac10.chm12004
f1_keywords:
- vbaac10.chm12004
ms.prod: access
api_name:
- Access.CustomControl.Controls
ms.assetid: 9e8e9948-94eb-87d3-6917-be95224da5c4
ms.date: 06/08/2017
---


# CustomControl.Controls Property (Access)

Returns the  **Controls** collection of a form, subform, report or section. Read-only **Controls**.


## Syntax

 _expression_. **Controls**

 _expression_ A variable that represents a **CustomControl** object.


## Remarks

Use the  **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax below returns the number of controls located on Form1. The second references the name of a property within a control.


```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```


## See also


#### Concepts


[CustomControl Object](Access.CustomControl.md)


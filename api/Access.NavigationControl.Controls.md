---
title: NavigationControl.Controls property (Access)
keywords: vbaac10.chm11036
f1_keywords:
- vbaac10.chm11036
ms.prod: access
api_name:
- Access.NavigationControl.Controls
ms.assetid: 68c6abcf-7bb7-4795-8c6c-685ed1c25dc9
ms.date: 02/21/2019
localization_priority: Normal
---


# NavigationControl.Controls property (Access)

Returns the **Controls** collection of a form, subform, report, or section. Read-only **Controls**.


## Syntax

_expression_.**Controls**

_expression_ A variable that represents a **[NavigationControl](Access.NavigationControl.md)** object.


## Remarks

Use the **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax returns the number of controls located on Form1. The second references the name of a property within a control.

```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
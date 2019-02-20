---
title: ComboBox.Controls property (Access)
keywords: vbaac10.chm11367
f1_keywords:
- vbaac10.chm11367
ms.prod: access
api_name:
- Access.ComboBox.Controls
ms.assetid: 8f936303-1d90-d1cd-320f-de175df686cf
ms.date: 02/21/2019
localization_priority: Normal
---


# ComboBox.Controls property (Access)

Returns the **Controls** collection of a form, subform, report, or section. Read-only **Controls**.


## Syntax

_expression_.**Controls**

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Remarks

Use the **Controls** property to refer to one of the controls on a form, subform, report, or section within or attached to another control. For example, the first code syntax returns the number of controls located on Form1. The second references the name of a property within a control.

```vb
Forms("Form1").Controls.Count 
 
Forms("Form1").Controls("Textbox1").Properties(5).Name
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
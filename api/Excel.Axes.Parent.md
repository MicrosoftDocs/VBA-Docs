---
title: Axes.Parent property (Excel)
keywords: vbaxl10.chm571075
f1_keywords:
- vbaxl10.chm571075
ms.prod: excel
api_name:
- Excel.Axes.Parent
ms.assetid: d5cd5daf-7579-4df3-8dad-b3daf3e5b5ae
ms.date: 04/13/2019
localization_priority: Normal
---


# Axes.Parent property (Excel)

Returns the parent object for the specified object. Read-only.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents an **[Axes](Excel.Axes(object).md)** object.


## Example

This example displays the name of the chart that contains _myAxis_.

```vb
Sub DisplayParentName() 
 
 Set myAxis = Charts(1).Axes(xlValue) 
 MsgBox myAxis.Parent.Name 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
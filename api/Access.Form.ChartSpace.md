---
title: Form.ChartSpace property (Access)
keywords: vbaac10.chm13522
f1_keywords:
- vbaac10.chm13522
ms.prod: access
api_name:
- Access.Form.ChartSpace
ms.assetid: e05f312f-d02b-bea5-7355-0a427834281c
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.ChartSpace property (Access)

Returns a **ChartSpace** object. Read-only.


## Syntax

_expression_.**ChartSpace**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

You must set a reference to the Microsoft Office web components type library to use this property.


## Example

This example reports the version of Microsoft Office web components in use for the specified form.


```vb
Dim objChartSpace As ChartSpace 
 
Set objChartSpace = Forms(0).ChartSpace 
 
MsgBox "Current version of Office Web Components: " _ 
 & objChartSpace.Version 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Application.DisplayFunctionToolTips property (Excel)
keywords: vbaxl10.chm133268
f1_keywords:
- vbaxl10.chm133268
ms.prod: excel
api_name:
- Excel.Application.DisplayFunctionToolTips
ms.assetid: cc294f6d-3e81-9fdc-b758-0a581b03ba9c
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DisplayFunctionToolTips property (Excel)

**True** if function ToolTips can be displayed. Read/write **Boolean**.


## Syntax

_expression_.**DisplayFunctionToolTips**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

In this example, Microsoft Excel notifies the user of the status of displaying function ToolTips.

```vb
Sub CheckToolTip() 
 
 ' Notify the user of the ability to display function ToolTips. 
 If Application.DisplayFunctionToolTips = True Then 
 MsgBox "The ability to display function ToolTips is on." 
 Else 
 MsgBox "The ability to display function ToolTips is off." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
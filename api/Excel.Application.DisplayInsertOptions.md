---
title: Application.DisplayInsertOptions property (Excel)
keywords: vbaxl10.chm133274
f1_keywords:
- vbaxl10.chm133274
ms.prod: excel
api_name:
- Excel.Application.DisplayInsertOptions
ms.assetid: 81c1d837-463f-bc33-f815-7c6dc9678d1b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DisplayInsertOptions property (Excel)

 **True** if the **Insert Options** button should be displayed. Read/write **Boolean**.


## Syntax

_expression_. `DisplayInsertOptions`

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

In this example, Microsoft Excel notifies the user the status of displaying the  **Insert Options** button.


```vb
Sub SettingToolTip() 
 
 ' Notify the user of the ToolTip status. 
 If Application.DisplayInsertOptions = True Then 
 MsgBox "The ability to display the Insert Options button is on." 
 Else 
 MsgBox "The ability to display the Insert Options button is off." 
 End If 
 
End Sub
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
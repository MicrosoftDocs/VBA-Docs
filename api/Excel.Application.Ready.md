---
title: Application.Ready property (Excel)
keywords: vbaxl10.chm133260
f1_keywords:
- vbaxl10.chm133260
api_name:
- Excel.Application.Ready
ms.assetid: 4b9577ee-0f0c-dd0b-c1dd-90cde2c5fb1e
ms.date: 04/05/2019
ms.localizationpriority: medium
---


# Application.Ready property (Excel)

Returns **True** when the Microsoft Excel application is ready; **False** when the Excel application is not ready. Read-only **Boolean**. 


## Syntax

_expression_.**Ready**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

In this example, Excel checks to see if the **Ready** property is set to **True**, and if so, a message displays "Application is ready." Otherwise, Excel displays the message "Application is not ready."

```vb
Sub UseReady() 
 
 If Application.Ready = True Then 
 MsgBox "Application is ready." 
 Else 
 MsgBox "Application is not ready." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
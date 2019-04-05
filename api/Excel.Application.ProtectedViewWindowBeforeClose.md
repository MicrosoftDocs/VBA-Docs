---
title: Application.ProtectedViewWindowBeforeClose event (Excel)
keywords: vbaxl10.chm504110
f1_keywords:
- vbaxl10.chm504110
ms.prod: excel
api_name:
- Excel.Application.ProtectedViewWindowBeforeClose
ms.assetid: 5fa37062-61c7-3002-1ea0-c5bd396b6a9b
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.ProtectedViewWindowBeforeClose event (Excel)

Occurs immediately before a Protected View window or a workbook in a Protected View window closes.


## Syntax

_expression_.**ProtectedViewWindowBeforeClose** (_Pvw_, _Reason_, _Cancel_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pvw_|Required| **[ProtectedViewWindow](Excel.ProtectedViewWindow.md)**|An object that represents the Protected View window that is closed.|
| _Reason_|Required| **[XlProtectedViewCloseReason](Excel.XlProtectedViewCloseReason.md)**|A constant that specifies the reason the Protected View window is closed.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the window does not close when the procedure is finished.|

## Return value

Nothing


## Example

The following code example prompts the user for a yes or no response before closing the Protected View window. This code must be placed in a class module, and an instance of that class must be correctly initialized. 

For more information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


```vb
Private Sub App_ProtectedViewWindowBeforeClose(ByVal Pvw as ProtectedViewWindow, _ 
 Reason as XlProtectedViewCloseReason, Cancel as Boolean) 
 a = MsgBox("Do you really want to close the Protected View window?", _ 
 vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
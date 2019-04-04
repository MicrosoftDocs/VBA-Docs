---
title: Application.ProtectedViewWindowBeforeEdit event (Excel)
keywords: vbaxl10.chm504109
f1_keywords:
- vbaxl10.chm504109
ms.prod: excel
api_name:
- Excel.Application.ProtectedViewWindowBeforeEdit
ms.assetid: b823b4a4-5d2f-7caf-f66f-5053b58082e4
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.ProtectedViewWindowBeforeEdit event (Excel)

Occurs immediately before editing is enabled on the workbook in the specified Protected View window.


## Syntax

_expression_.**ProtectedViewWindowBeforeEdit** (_Pvw_, _Cancel_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pvw_|Required| **[ProtectedViewWindow](Excel.ProtectedViewWindow.md)**|The Protected View window that contains the workbook that is enabled for editing.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, editing is not enabled on the workbook.|

## Return value

Nothing


## Example

The following code example prompts the user for a yes or no response before enabling editing on a workbook in a Protected View window. This code must be placed in a class module, and an instance of the class must be correctly initialized. 

For more information about how to use event procedures with the **Application** object, see [Using events with the Application object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


```vb
Private Sub App_ProtectedViewWindowBeforeEdit(ByVal Pvw As ProtectedViewWindow, Cancel As Boolean) 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Do you really " _ 
 & "want to edit the workbook?", _ 
 vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
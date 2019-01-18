---
title: Application.WorkbookBeforeClose Event (Excel)
keywords: vbaxl10.chm504084
f1_keywords:
- vbaxl10.chm504084
ms.prod: excel
api_name:
- Excel.Application.WorkbookBeforeClose
ms.assetid: 9c3618ea-0e5e-e4fe-20af-279826bfa7c3
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WorkbookBeforeClose Event (Excel)

Occurs immediately before any open workbook closes.


## Syntax

_expression_. `WorkbookBeforeClose`( `_Wb_` , `_Cancel_` )

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The workbook that's being closed|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the workbook doesn't close when the procedure is finished.|

## Return value

Nothing


## Example

This example prompts the user for a yes or no response before closing any workbook. For more information about how to use event procedures with the  **Application** object, see [Using Events with the Application Object](../excel/Concepts/Events-WorksheetFunctions-Shapes/using-events-with-the-application-object.md).


```vb
Private Sub App_WorkbookBeforeClose(ByVal Wb as Workbook, _ 
 Cancel as Boolean) 
 a = MsgBox("Do you really want to close the workbook?", _ 
 vbYesNo) 
 If a = vbNo Then Cancel = True 
End Sub
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
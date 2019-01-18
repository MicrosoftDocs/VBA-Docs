---
title: Window.Close method (Excel)
keywords: vbaxl10.chm356081
f1_keywords:
- vbaxl10.chm356081
ms.prod: excel
api_name:
- Excel.Window.Close
ms.assetid: 59772b1d-823e-70d6-97d7-a5ddb44abdcc
ms.date: 06/08/2017
localization_priority: Priority
---


# Window.Close method (Excel)

Closes the object.


## Syntax

_expression_. `Close`( `_SaveChanges_` , `_Filename_` , `_RouteWorkbook_` )

_expression_ A variable that represents a [Window](./Excel.Window.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Variant**|If there are no changes to the workbook, this argument is ignored. If there are changes to the workbook and the workbook appears in other open windows, this argument is ignored. If there are changes to the workbook but the workbook doesn't appear in any other open windows, this argument specifies whether changes should be saved. If set to  **True** , changes are saved to the workbook. If there is not yet a file name associated with the workbook, then _Filename_ is used. If _Filename_ is omitted, the user is asked to supply a file name.|
| _Filename_|Optional| **Variant**|Save changes under this file name.|
| _RouteWorkbook_|Optional| **Variant**|If the workbook doesn't need to be routed to the next recipient (if it has no routing slip or has already been routed), this argument is ignored. Otherwise, Microsoft Excel routes the workbook according to the value of this parameter. If set to  **True** , the workbook is sent to the next recipient. If set to **False** , the workbook is not sent. If omitted, the user is asked whether the workbook should be sent.|

## Return value

 **True** if the method successfully closed the object; otherwise, **False**.


## Remarks

Closing a workbook from Visual Basic doesn't run any Auto_Close macros in the workbook. Use the  **[RunAutoMacros](Excel.Workbook.RunAutoMacros.md)** method to run the auto close macros.


## See also


[Window Object](Excel.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
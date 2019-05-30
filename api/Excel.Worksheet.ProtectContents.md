---
title: Worksheet.ProtectContents property (Excel)
keywords: vbaxl10.chm174090
f1_keywords:
- vbaxl10.chm174090
ms.prod: excel
api_name:
- Excel.Worksheet.ProtectContents
ms.assetid: 807717f6-1265-2d5d-5221-bc46b24d8281
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.ProtectContents property (Excel)

**True** if the contents of the sheet are protected. This protects the individual cells. To turn on content protection, use the **[Protect](Excel.Worksheet.Protect.md)** method with the _Contents_ argument set to **True**. Read-only **Boolean**.


## Syntax

_expression_.**ProtectContents**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example displays a message box if the contents of Sheet1 are protected.

```vb
If Worksheets("Sheet1").ProtectContents = True Then 
 MsgBox "The contents of Sheet1 are protected." 
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
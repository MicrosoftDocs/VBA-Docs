---
title: Worksheet.Delete method (Excel)
keywords: vbaxl10.chm174075
f1_keywords:
- vbaxl10.chm174075
ms.prod: excel
api_name:
- Excel.Worksheet.Delete
ms.assetid: a51e1673-e09d-824f-1acc-dda18c120204
ms.date: 08/24/2018
localization_priority: Normal
---


# Worksheet.Delete method (Excel)

Deletes the object.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Return value

Boolean


## Remarks

When you delete a **[Worksheet](Excel.Worksheet.md)**, this method displays a dialog box that prompts the user to confirm the deletion. This dialog box is displayed by default. When called on the **Worksheet** object, the **Delete** method returns a **Boolean** value that is **False** if the user clicked **Cancel** on the dialog box or **True** if the user clicked **Delete**.

To delete a worksheet without displaying a dialog box, set the **[Application.DisplayAlerts](Excel.Application.DisplayAlerts.md)** property to **False**.

## See also

- [Worksheet Object](Excel.Worksheet.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

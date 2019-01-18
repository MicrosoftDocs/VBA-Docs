---
title: Using events with Excel objects
keywords: vbaxl10.chm5205782
f1_keywords:
- vbaxl10.chm5205782
ms.prod: excel
ms.assetid: f5fac10f-17f4-2c8c-f39f-c2b616c8e895
ms.date: 11/13/2018
localization_priority: Normal
---


# Using events with Excel objects

You can write event procedures in Microsoft Excel at the worksheet, chart, query table, workbook, or application level. For example, the **Activate** event occurs at the sheet level, and the **SheetActivate** event is available at both the workbook and application levels. The **SheetActivate** event for a workbook occurs when any sheet in the workbook is activated. At the application level, the **SheetActivate** event occurs when any sheet in any open workbook is activated.

[Worksheet](worksheet-object-events.md), [Chart](chart-object-events.md), and event procedures are available for any open sheet or workbook. To write event procedures for an [embedded chart](chart-object-events.md), a **[QueryTable](querytable-object-events.md)** object, or an **[Application](../../../api/Excel.Application(object).md)** object, you must create a new object by using the **WithEvents** keyword in a class module.

Use the **EnableEvents** property to enable or disable events. For example, using the **Save** method to save a workbook causes the BeforeSave event to occur. You can prevent this by setting the **EnableEvents** property to **False** before you call the **Save** method.


## Example

```vb
Application.EnableEvents = False 
ActiveWorkbook.Save 
Application.EnableEvents = True
```

## See also

- [Excel functions (by category)](https://support.office.com/article/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
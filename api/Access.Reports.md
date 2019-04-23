---
title: Reports object (Access)
keywords: vbaac10.chm12478
f1_keywords:
- vbaac10.chm12478
ms.prod: access
api_name:
- Access.Reports
ms.assetid: 37c5f55e-3c3a-6140-d305-7e8118d9d2b1
ms.date: 03/21/2019
localization_priority: Normal
---


# Reports object (Access)

The **Reports** collection contains all of the currently open reports in a Microsoft Access database.


## Remarks

You can use the **Reports** collection in Visual Basic or in an expression to refer to reports that are currently open. For example, you can enumerate the **Reports** collection to set or return the values of properties of individual reports in the collection.

You can refer to an individual **[Report](Access.Report.md)** object in the **Reports** collection either by referring to the report by name, or by referring to its index within the collection.

The **Reports** collection is indexed beginning with zero. If you refer to a report by its index, the first report is Reports(0), the second report is Reports(1), and so on. If you opened Report1 and then opened Report2, Report2 would be referenced in the **Reports** collection by its index as Reports(1). If you then closed Report1, Report2 would be referenced in the **Reports** collection by its index as Reports(0).

> [!NOTE] 
> To list all reports in the database, whether open or closed, enumerate the **[AllReports](Access.AllReports.md)** collection of the **[CurrentProject](Access.CurrentProject.md)** object. You can then use the **Name** property of each individual **[AccessObject](Access.AccessObject.md)** object to return the name of a report.

You can't add or delete a **Report** object from the **Reports** collection.


## Properties

- [Application](Access.Reports.Application.md)
- [Count](Access.Reports.Count.md)
- [Item](Access.Reports.Item.md)
- [Parent](Access.Reports.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: ListObject.Refresh method (Excel)
keywords: vbaxl10.chm734075
f1_keywords:
- vbaxl10.chm734075
ms.prod: excel
api_name:
- Excel.ListObject.Refresh
ms.assetid: 7827a116-0ba4-9855-e0e9-550a85d36ed3
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.Refresh method (Excel)

Retrieves the current data and schema for the list from the server that is running Microsoft SharePoint Foundation. This method can be used only with lists that are linked to a SharePoint site. If the SharePoint site is not available, calling this method returns an error.


## Syntax

_expression_.**Refresh**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Remarks

Calling the **Refresh** method does not commit changes to the list in the Excel workbook. Uncommitted changes in the list in Excel are discarded when the **Refresh** method is called.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

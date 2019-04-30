---
title: ListObject.Delete method (Excel)
keywords: vbaxl10.chm734073
f1_keywords:
- vbaxl10.chm734073
ms.prod: excel
api_name:
- Excel.ListObject.Delete
ms.assetid: cd621c14-5e13-b51b-2b39-29118aeac3c8
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.Delete method (Excel)

Deletes the **ListObject** object and clears the cell data from the worksheet.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Remarks

If the list is linked to a SharePoint site, deleting it does not affect data on the server that is running SharePoint Foundation. Any uncommitted changes made to the local list are not sent to the SharePoint list. There is no warning that these uncommitted changes are lost.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
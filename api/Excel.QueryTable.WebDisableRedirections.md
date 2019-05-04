---
title: QueryTable.WebDisableRedirections property (Excel)
keywords: vbaxl10.chm518129
f1_keywords:
- vbaxl10.chm518129
ms.prod: excel
api_name:
- Excel.QueryTable.WebDisableRedirections
ms.assetid: 36aec986-de9c-2c7e-a07c-ae77d75d4c7c
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.WebDisableRedirections property (Excel)

**True** if web query redirections are disabled for a **QueryTable** object. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**WebDisableRedirections**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **WebDisableRedirections** property applies only to **QueryTable** objects.


## Example

In this example, Microsoft Excel determines the settings of web query redirections for the first worksheet in the workbook. This example assumes that a **QueryTable** object exists on the first worksheet; otherwise, a run-time error occurs.

```vb
Sub CheckWebQuerySetting() 
 Dim wksSheet As Worksheet 
 Set wksSheet = Application.ActiveSheet 
 MsgBox wksSheet.QueryTables(1).WebDisableRedirections 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
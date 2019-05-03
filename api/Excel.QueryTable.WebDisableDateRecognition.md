---
title: QueryTable.WebDisableDateRecognition property (Excel)
keywords: vbaxl10.chm518127
f1_keywords:
- vbaxl10.chm518127
ms.prod: excel
api_name:
- Excel.QueryTable.WebDisableDateRecognition
ms.assetid: 6db374e2-67b2-bf84-35d4-dd87494c0d81
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.WebDisableDateRecognition property (Excel)

**True** if data that resembles dates is parsed as text when you import a webpage into a query table. **False** if date recognition is used. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**WebDisableDateRecognition**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Use this property only when the query table's **[QueryType](Excel.QueryTable.QueryType.md)** property is set to **xlWebQuery** and the query returns an HTML document.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **WebDisableDateRecognition** property applies only to **QueryTable** objects.


## Example

This example turns off date recognition so that webpage data that resembles dates is imported as text, and then refreshes the query table.

```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "URL;https://datasvr/98q1/19980331.htm", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .WebDisableDateRecognition = True 
 .Refresh 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
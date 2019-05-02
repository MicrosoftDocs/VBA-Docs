---
title: QueryTable.PostText property (Excel)
keywords: vbaxl10.chm518089
f1_keywords:
- vbaxl10.chm518089
ms.prod: excel
api_name:
- Excel.QueryTable.PostText
ms.assetid: f89c21bb-2b51-49b2-b986-8c3aca2038c1
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.PostText property (Excel)

Returns or sets the string used with the post method of inputting data into a web server to return data from a web query. Read/write **String**.


## Syntax

_expression_.**PostText**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Microsoft Excel includes sample web queries that you can modify by changing the HTML code by using WordPad or another text editor. You can find these samples in the Queries folder where you installed Microsoft Office.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **PostText** property applies only to **QueryTable** objects.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
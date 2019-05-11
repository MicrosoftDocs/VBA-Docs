---
title: Range.ID property (Excel)
keywords: vbaxl10.chm144231
f1_keywords:
- vbaxl10.chm144231
ms.prod: excel
api_name:
- Excel.Range.ID
ms.assetid: 0ff7f261-8829-2858-5097-a638c01e5f3c
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.ID property (Excel)

Returns or sets a **String** value that represents the identifying label for the specified cell when the page is saved as a webpage.


## Syntax

_expression_.**ID**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

You can use an ID label as a hyperlink reference in other HTML documents or on the same webpage.


## Example

This example sets the ID of cell A1 on the active worksheet to Target.

```vb
ActiveSheet.Range("A1").ID = "target"
```

<br/>

Later, the document is saved as a webpage, and the following line of HTML is added to the webpage. When the user then views the page in a web browser and chooses the hyperlink, the browser displays the cell.

```vb
<A HREF="#target">Quarterly earnings</A>
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: PublishObject.Title property (Excel)
keywords: vbaxl10.chm652080
f1_keywords:
- vbaxl10.chm652080
ms.prod: excel
api_name:
- Excel.PublishObject.Title
ms.assetid: 3e8eae5c-62f5-3d72-2c27-ff5107153adc
ms.date: 06/08/2017
localization_priority: Normal
---


# PublishObject.Title property (Excel)

Returns or sets the title of the webpage when the document is saved as a webpage. Read/write  **String**.


## Syntax

_expression_.**Title**

_expression_ A variable that represents a **[PublishObject](Excel.PublishObject.md)** object.


## Remarks

The title is usually displayed in the window title bar when the document is viewed in the web browser.


## Example

This example sets the webpage title to "Sales Forecast" when the first item in the first workbook is saved as a webpage.


```vb
Workbooks(1).PublishObjects(1).Title = "Sales Forecast"
```


## See also


[PublishObject Object](Excel.PublishObject.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
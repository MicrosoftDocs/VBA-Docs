---
title: Style.IncludeAlignment property (Excel)
keywords: vbaxl10.chm177080
f1_keywords:
- vbaxl10.chm177080
ms.prod: excel
api_name:
- Excel.Style.IncludeAlignment
ms.assetid: 4b58251d-cf1f-3301-a597-3e2c756144fe
ms.date: 05/16/2019
localization_priority: Normal
---


# Style.IncludeAlignment property (Excel)

**True** if the style includes the **AddIndent**, **HorizontalAlignment**, **VerticalAlignment**, **WrapText**, **IndentLevel**, and **Orientation** properties of the **Style** object. Read/write **Boolean**.


## Syntax

_expression_.**IncludeAlignment**

_expression_ A variable that represents a **[Style](Excel.Style.md)** object.


## Example

This example sets the style attached to cell A1 on Sheet1 to include alignment format.

```vb
Worksheets("Sheet1").Range("A1").Style.IncludeAlignment = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
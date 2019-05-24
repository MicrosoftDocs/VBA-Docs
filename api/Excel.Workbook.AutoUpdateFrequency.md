---
title: Workbook.AutoUpdateFrequency property (Excel)
keywords: vbaxl10.chm199078
f1_keywords:
- vbaxl10.chm199078
ms.prod: excel
api_name:
- Excel.Workbook.AutoUpdateFrequency
ms.assetid: dfded5c8-94d6-8a0f-24c1-248bd502850b
ms.date: 05/25/2019
localization_priority: Normal
---


# Workbook.AutoUpdateFrequency property (Excel)

Returns or sets the number of minutes between automatic updates to the shared workbook. Read/write **Long**.


## Syntax

_expression_.**AutoUpdateFrequency**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

The **AutoUpdateFrequency** property must be set to a value from 5 to 1440 for this property to take effect.


## Example

This example causes the shared workbook to be automatically updated every five minutes.

```vb
ActiveWorkbook.AutoUpdateFrequency = 5
```

> [!NOTE] 
> Workbook sharing must be enabled or you may see the following error: "Method 'AutoUpdateFrequency' of object '_Workbook' failed".




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
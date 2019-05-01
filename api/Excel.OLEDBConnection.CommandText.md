---
title: OLEDBConnection.CommandText property (Excel)
keywords: vbaxl10.chm794076
f1_keywords:
- vbaxl10.chm794076
ms.prod: excel
api_name:
- Excel.OLEDBConnection.CommandText
ms.assetid: 2c5e976c-513f-24b0-f25e-056fc9babaf9
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBConnection.CommandText property (Excel)

Returns or sets the command string for the specified data source. Read/write **Variant**.


## Syntax

_expression_.**CommandText**

_expression_ A variable that represents an **[OLEDBConnection](Excel.OLEDBConnection.md)** object.


## Remarks

You should use the **CommandText** property instead of the **SQL** property, which now exists primarily for compatibility with earlier versions of Microsoft Excel. If you use both properties, the **CommandText** property's value takes precedence.

The **[CommandType](Excel.OLEDBConnection.CommandType.md)** property describes the value of the **CommandText** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Workbook.PersonalViewListSettings property (Excel)
keywords: vbaxl10.chm199122
f1_keywords:
- vbaxl10.chm199122
ms.prod: excel
api_name:
- Excel.Workbook.PersonalViewListSettings
ms.assetid: 998320bf-d703-e42f-8b43-5a7b909a846d
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.PersonalViewListSettings property (Excel)

**True** if filter and sort settings for lists are included in the user's personal view of the shared workbook. Read/write **Boolean**.


## Syntax

_expression_.**PersonalViewListSettings**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example removes print settings and filter and sort settings from the user's personal view of workbook two.

```vb
With Workbooks(2) 
 .PersonalViewListSettings = False 
 .PersonalViewPrintSettings = False 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
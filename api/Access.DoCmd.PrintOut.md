---
title: DoCmd.PrintOut method (Access)
keywords: vbaac10.chm4166
f1_keywords:
- vbaac10.chm4166
ms.prod: access
api_name:
- Access.DoCmd.PrintOut
ms.assetid: 3b7c1ab7-1a60-cab3-2d4e-c95d6b5bd4aa
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.PrintOut method (Access)

The **PrintOut** method carries out the PrintOut action in Visual Basic.


## Syntax

_expression_.**PrintOut** (_PrintRange_, _PageFrom_, _PageTo_, _PrintQuality_, _Copies_, _CollateCopies_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PrintRange_|Optional|**[AcPrintRange](Access.AcPrintRange.md)**|An **AcPrintRange** constant that specifies the range to print. The default value is **acPrintAll**.|
| _PageFrom_|Optional|**Variant**|The first page to print. A numeric expression that's a valid page number in the active form or datasheet. This argument is required if you specify **acPages** for the _PrintRange_ argument.|
| _PageTo_|Optional|**Variant**|The last page to print. A numeric expression that's a valid page number in the active form or datasheet. This argument is required if you specify **acPages** for the _PrintRange_ argument.|
| _PrintQuality_|Optional|**[AcPrintQuality](Access.AcPrintQuality.md)**|An **AcPrintQuality** constant that specifies the print quality. The default value is **acHigh**.|
| _Copies_|Optional|**Variant**|The number of copies to print. If you leave this argument blank, the default (1) is assumed.|
| _CollateCopies_|Optional|**Variant**|Use **True** (1) to collate copies and **False** (0) to print without collating. If you leave this argument blank, the default (**True**) is assumed.|

## Remarks

You can use the PrintOut action to print the active object in the open database. You can print datasheets, reports, forms, data access pages, and modules.


## Example

The following example prints two collated copies of the first four pages of the active form or datasheet.

```vb
DoCmd.PrintOut acPages, 1, 4, , 2
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

---
title: ColumnFormat.FieldFormat property (Outlook)
keywords: vbaol11.chm2729
f1_keywords:
- vbaol11.chm2729
ms.prod: outlook
api_name:
- Outlook.ColumnFormat.FieldFormat
ms.assetid: 14064b56-65c2-1c7d-1e74-3bfa2d2ccaa7
ms.date: 06/08/2017
localization_priority: Normal
---


# ColumnFormat.FieldFormat property (Outlook)

Returns or sets a **Long** value that represents the display format of the property to which the **[ColumnFormat](Outlook.ColumnFormat.md)** object is associated. Read/write.


## Syntax

_expression_. `FieldFormat`

_expression_ A variable that represents a [ColumnFormat](Outlook.ColumnFormat.md) object.


## Remarks

The value of this property is a constant from an enumeration, where the enumeration is dependent on the value of the  **[FieldType](Outlook.ColumnFormat.FieldType.md)** property for the **ColumnFormat** object:



| **FieldType value**| **FieldFormat enumeration**|
| **olCurrency**| **[OlFormatCurrency](Outlook.OlFormatCurrency.md)**|
| **olFormatDateTime**| **[OlFormatDateTime](Outlook.OlFormatDateTime.md)**|
| **olDuration**| **[OlFormatDuration](Outlook.OlFormatDuration.md)**|
| **olInteger**| **[OlFormatInteger](Outlook.OlFormatInteger.md)**|
| **olKeywords**| **[OlFormatKeywords](Outlook.OlFormatKeywords.md)**|
| **olNumber**| **[OlFormatNumber](Outlook.OlFormatNumber.md)**|
| **olPercent**| **[OlFormatPercent](Outlook.OlFormatPercent.md)**|
| **olText**| **[OlFormatText](Outlook.OlFormatText.md)**|
| **olYesNo**| **[OlFormatYesNo](Outlook.OlFormatYesNo.md)**|
| **olEnumeration**| **[OlFormatEnumeration](Outlook.OlFormatEnumeration.md)**|
| **olSmartFrom**| **[OlFormatSmartFrom](Outlook.OlFormatSmartFrom.md)**|

For field types not listed in the above table, the value of this property is ignored.


## See also


[ColumnFormat Object](Outlook.ColumnFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
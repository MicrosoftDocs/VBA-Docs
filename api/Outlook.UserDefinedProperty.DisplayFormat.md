---
title: UserDefinedProperty.DisplayFormat property (Outlook)
keywords: vbaol11.chm8
f1_keywords:
- vbaol11.chm8
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperty.DisplayFormat
ms.assetid: f891aa8d-a769-275d-c027-7c5260eafc97
ms.date: 06/08/2017
localization_priority: Normal
---


# UserDefinedProperty.DisplayFormat property (Outlook)

Returns a **Long** value that represents the display format for the **[UserDefinedProperty](Outlook.UserDefinedProperty.md)** object. Read-only.


## Syntax

_expression_.**DisplayFormat**

_expression_ A variable that represents a [UserDefinedProperty](Outlook.UserDefinedProperty.md) object.


## Remarks

The value of this property is a constant from an enumeration, where the enumeration is dependent on the value of the  **[Type](Outlook.UserDefinedProperty.Type.md)** property for the **UserDefinedProperty** object:



| **Type value**| **DisplayFormat enumeration**|
| **olCombination**|No enumeration available. This property always returns 1 for  **olCombination**.|
| **olCurrency**| **[OlFormatCurrency](Outlook.OlFormatCurrency.md)**|
| **olDateTime**| **[OlFormatDateTime](Outlook.OlFormatDateTime.md)**|
| **olDuration**| **[OlFormatDuration](Outlook.OlFormatDuration.md)**|
| **olEnumeration**| **[OlFormatEnumeration](Outlook.OlFormatEnumeration.md)**|
| **olFormula**|No enumeration available. This property always returns 1 for  **olFormula**.|
| **olInteger**| **[OlFormatInteger](Outlook.OlFormatInteger.md)**|
| **olKeywords**| **[OlFormatKeywords](Outlook.OlFormatKeywords.md)**|
| **olNumber**| **[OlFormatNumber](Outlook.OlFormatNumber.md)**|
| **olOutlookInternal**|No enumeration available. This property always returns 1 for  **olOutlookInternal**.|
| **olPercent**| **[OlFormatPercent](Outlook.OlFormatPercent.md)**|
| **olSmartFrom**| **[OlFormatSmartFrom](Outlook.OlFormatSmartFrom.md)**|
| **olText**| **[OlFormatText](Outlook.OlFormatText.md)**|
| **olYesNo**| **[OlFormatYesNo](Outlook.OlFormatYesNo.md)**|

## See also


[UserDefinedProperty Object](Outlook.UserDefinedProperty.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
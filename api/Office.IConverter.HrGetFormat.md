---
title: IConverter.HrGetFormat method (Office)
keywords: vbaof11.chm333005
f1_keywords:
- vbaof11.chm333005
ms.prod: office
api_name:
- Office.IConverter.HrGetFormat
ms.assetid: bdee0e57-d87b-f1ec-950e-d8b676fd91db
ms.date: 01/16/2019
localization_priority: Normal
---


# IConverter.HrGetFormat method (Office)

This feature is only available in the [Open XML Format SDK](https://docs.microsoft.com/office/open-xml/open-xml-sdk).


## Syntax

_expression_.**HrGetFormat** (_bstrPath_, _pbstrClass_, _pcap_, _ppcp_, _pcuic_)

_expression_ An expression that returns an **[IConverter](Office.IConverter.md)** object.


## Parameters

|Name|Required/Optional|Data type|
|:---|:----------------|:--------|
| _bstrPath_|Required|**String**|
| _pbstrClass_|Required|**String**|
| _pcap_|Required|**IConverterApplicationPreferences**|
| _ppcp_|Required|**IConverterPreferences**|
| _pcuic_|Required|**IConverterUICallback**|

## Return value

[HRESULT]


## See also

- [IConverter object members](overview/Library-Reference/iconverter-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: IConverter.HrImport method (Office)
keywords: vbaof11.chm333003
f1_keywords:
- vbaof11.chm333003
ms.prod: office
api_name:
- Office.IConverter.HrImport
ms.assetid: cc329008-286a-81aa-d9c6-b5cd3bf05e37
ms.date: 01/16/2019
localization_priority: Normal
---


# IConverter.HrImport method (Office)

This feature is only available in the [Open XML Format SDK](https://docs.microsoft.com/office/open-xml/open-xml-sdk).


## Syntax

_expression_.**HrImport** (_bstrSourcePath_, _bstrDestPath_, _pcap_, _ppcp_, _pcuic_)

_expression_ An expression that returns an **[IConverter](Office.IConverter.md)** object.


## Parameters

|Name|Required/Optional|Data type|
|:---|:----------------|:--------|
| _bstrSourcePath_|Required|**String**|
| _bstrDestPath_|Required|**String**|
| _pcap_|Required|**IConverterApplicationPreferences**|
| _ppcp_|Required|**IConverterPreferences**|
| _pcuic_|Required|**IConverterUICallback**|

## Return value

[HRESULT]


## See also

- [IConverter object members](overview/Library-Reference/iconverter-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Range.ConvertToLinkedDataType method (Excel)
keywords: vbaxl10.chm144263
f1_keywords:
- vbaxl10.chm144263
ms.prod: excel
api_name:
- Excel.Range.ConvertToLinkedDataType
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.ConvertToLinkedDataType method (Excel)

Attempts to convert all the cells in the range to a Linked data type such as [Stocks or Geography](https://support.office.com/article/stock-quotes-and-geographic-data-61a33056-9935-484f-8ac8-f1a89e210877).

## Syntax

_expression_.**ConvertToLinkedDataType** (_ServiceID_, _LanguageCulture_)

_expression_ A variable that represents a **[Range](Excel.Range(Object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ServiceID_|Required| **Long**|The ID of the service that will provide the linked entity.|
| _LanguageCulture_|Required| **String**|A string representing the [LCID](https://docs.microsoft.com/openspecs/windows_protocols/ms-lcid/a9eac961-e77d-41a6-90a5-ce1a8b0cdb9c) of the language and culture that you would like to use for the linked entity. |

## Remarks

The method will fail and throw a runtime exception 1004 if the specified locale is not supported on the specified service.

It will have no effect (and throw no exception) in these cases:

- The cells in the range are blank (that is, there is nothing to convert).
- The cells in the range contain a formula. If you want to convert such a range, you need to set the cell values to the current calc result first.
- The cells in the range have already been converted to the specified data type.

## Example

This code will convert cell E5 to a _Stocks_ Linked data type in the US-English locale.

```vb
Range("E5").ConvertToLinkedDataType ServiceID:=268435456, LanguageCulture:= "en-US"
```

<br/>

This code will convert cell E6 to a _Geography_ Linked data type in the US-English locale.

```vb
Range("E6").ConvertToLinkedDataType ServiceID:=536870912, LanguageCulture:= "en-US"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

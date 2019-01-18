---
title: SetText method (Microsoft Forms)
keywords: fm20.chm2012330
f1_keywords:
- fm20.chm2012330
ms.prod: office
api_name:
- Office.SetText
ms.assetid: e7a246fb-eb50-7c35-1b9f-3e927589aa37
ms.date: 11/15/2018
localization_priority: Normal
---


# SetText method (Microsoft Forms)

Copies a text string to the **[DataObject](dataobject-object.md)** using a specified format.

## Syntax

_object_. **SetText(**_StoreData_ [, _format_ ] **)**

The **SetText** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _StoreData_|Required. Defines the data to store on the **DataObject**.|
| _format_|Optional. An integer or string specifying the format of _StoreData_. When retrieving data from the **DataObject**, the format identifies the piece of data to retrieve.|

## Settings

The settings for _format_ are:

|Value|Description|
|:-----|:-----|
|1|Text format.|
|A string or integer value other than 1|A user-defined **DataObject** format.|

## Remarks

The **DataObject** stores data according to its format. When the user supplies a string, the **DataObject** saves the text under the specified format.

If the **DataObject** contains data in the same format as new data, the new data replaces the existing data in the **DataObject**. If the new data is in a new format, the new data and the new format are both added to the **DataObject**, and the previously existing data is there as well.

If no format is specified, the **SetText** method assigns the Text format to the text string. If a new format is specified, the **DataObject** registers the new format with the system.

## See also

- [Standard Clipboard formats](https://docs.microsoft.com/windows/desktop/dataxchg/standard-clipboard-formats)
- [Registered Clipboard formats](https://docs.microsoft.com/windows/desktop/dataxchg/clipboard-formats)
- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
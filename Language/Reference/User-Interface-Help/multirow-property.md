---
title: MultiRow property 
keywords: fm20.chm5225068
f1_keywords:
- fm20.chm5225068
ms.prod: office
api_name:
- Office.MultiRow
ms.assetid: 2030addd-5a90-e94f-9647-a4aa50e68690
ms.date: 11/16/2018
localization_priority: Normal
---


# MultiRow property 

Specifies whether the control has more than one row of tabs.

## Syntax

_object_.**MultiRow** [= _Boolean_ ]

The **MultiRow** property syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the control has more than one row of tabs.|

## Settings

The settings for _Boolean_ are:

|Value|Description|
|:-----|:-----|
|**True**|Allows more than one row of tabs.|
|**False**|Restricts tabs to a single row (default).|

## Remarks

The width and number of tabs determines the number of rows. Changing the control's size also changes the number of rows. This allows the developer to resize the control and ensure that tabs wrap to fit the control. 

If the **MultiRow** property is **False**, truncation occurs if the width of the tabs exceeds the width of the control.

If **MultiRow** is **False** and tabs are truncated, there will be a small scroll bar on the **[TabStrip](tabstrip-control.md)** to allow scrolling to the other tabs or pages.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
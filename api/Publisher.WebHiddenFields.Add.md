---
title: WebHiddenFields.Add method (Publisher)
keywords: vbapb10.chm3997700
f1_keywords:
- vbapb10.chm3997700
ms.prod: publisher
api_name:
- Publisher.WebHiddenFields.Add
ms.assetid: c3035138-f369-b561-b1f8-9977bd9e080c
ms.date: 06/18/2019
localization_priority: Normal
---


# WebHiddenFields.Add method (Publisher)

Adds a new hidden field to a web form and returns a **Long** indicating the number of the new field in the **WebHiddenFields** collection. New fields are always placed at the end of the current field list.


## Syntax

_expression_.**Add** (_Name_, _Value_)

_expression_ A variable that represents a **[WebHiddenFields](Publisher.WebHiddenFields.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Name_|Required| **String**|The name of the new field.|
|_Value_|Required| **String**|The value of the new field.|

## Return value

Long


## Example

The following example adds a new hidden field to the specified web command button control. Shape one on page one of the active publication must be a web command button control for this example to work.

```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .WebCommandButton.HiddenFields _ 
 .Add Name:="subject", Value:="service request"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
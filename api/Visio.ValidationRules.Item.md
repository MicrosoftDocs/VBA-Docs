---
title: ValidationRules.Item property (Visio)
keywords: vis_sdr.chm18313765
f1_keywords:
- vis_sdr.chm18313765
ms.prod: visio
api_name:
- Visio.ValidationRules.Item
ms.assetid: 4133f9ba-ca20-104a-5a30-7de37b978706
ms.date: 06/08/2017
localization_priority: Normal
---


# ValidationRules.Item property (Visio)

Returns the  **[ValidationRule](Visio.ValidationRule.md)** object that has the specified index position. The **Item** property is the default property for all collections. Read-only.


## Syntax

_expression_.**Item** (_NameUIDOrIndex_)

_expression_ A variable that represents a **[ValidationRules](Visio.ValidationRules.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NameUOrIndex_|Required| **Variant**|The index number of the object in its collection.|

## Return value

 **ValidationRule**


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```vb
objRet = object(index)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
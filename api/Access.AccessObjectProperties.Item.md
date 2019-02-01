---
title: AccessObjectProperties.Item property (Access)
keywords: vbaac10.chm12701
f1_keywords:
- vbaac10.chm12701
ms.prod: access
api_name:
- Access.AccessObjectProperties.Item
ms.assetid: d26e6417-f7ec-3ba9-096f-1bbe5b354347
ms.date: 02/01/2019
localization_priority: Normal
---


# AccessObjectProperties.Item property (Access)

The **Item** property returns a specific member of a collection either by position or by index. Read-only **AccessObjectProperty**.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an **[AccessObjectProperties](Access.AccessObjectProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|An expression that specifies the position of a member of the collection referred to by the _expression_ argument.<br/><br/>If a numeric expression, the _index_ argument must be a number from 0 to the value of the collection's **Count** property minus 1.<br/><br/>If a string expression, the _index_ argument must be the name of a member of the collection.|

## Remarks

If the value provided for the _index_ argument doesn't match any existing member of the collection, an error occurs.

The **Item** property is the default member of a collection, so you don't have to specify it explicitly. For example, the following two lines of code are equivalent.

```vb
Debug.Print Modules(0)
```


```vb
Debug.Print Modules.Item(0)
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
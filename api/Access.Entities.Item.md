---
title: Entities.Item property (Access)
keywords: vbaac10.chm14562
f1_keywords:
- vbaac10.chm14562
ms.prod: access
api_name:
- Access.Entities.Item
ms.assetid: 6e8e9b66-35c9-d436-6391-df424ad0f66f
ms.date: 03/06/2019
localization_priority: Normal
---


# Entities.Item property (Access)

The **Item** property returns a specific member of a collection either by position or by index. Read-only **Object**.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an **[Entities](Access.Entities.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|An expression that specifies the position of a member of the collection referred to by the _expression_ argument.<br/><br/>If a numeric expression, the _Index_ argument must be a number from 0 to the value of the collection's **Count** property minus 1.<br/><br/>If a string expression, the _Index_ argument must be the name of a member of the collection.|

## Remarks

If the value provided for the _Index_ argument doesn't match any existing member of the collection, an error occurs.

The **Item** property is the default member of a collection, so you don't have to specify it explicitly. For example, the following two lines of code are equivalent.

```vb
Debug.Print Modules(0)
```

```vb
Debug.Print Modules.Item(0)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
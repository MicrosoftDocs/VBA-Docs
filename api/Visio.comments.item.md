---
title: Comments.Item property (Visio)
ms.prod: visio
ms.assetid: fed2a079-de87-d5ce-1d74-0bfa5a328441
ms.date: 06/08/2017
localization_priority: Normal
---


# Comments.Item property (Visio)

Returns an object from a collection. The  **Item** property is the default property for all collections. Read-only[Comment](Visio.comment.md).


## Syntax

_expression_.**Item**

_expression_ A variable that represents a **[Comments](Visio.Comments.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|Contains the index of the object to retrieve.|

## Return value

 **Comment**


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```vb
objRet = object(index )
```


## Property value

 **IVCOMMENT**


## See also


[Comments Collection](Visio.comments.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
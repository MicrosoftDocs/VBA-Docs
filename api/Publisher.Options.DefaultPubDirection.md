---
title: Options.DefaultPubDirection property (Publisher)
keywords: vbapb10.chm1048624
f1_keywords:
- vbapb10.chm1048624
ms.prod: publisher
api_name:
- Publisher.Options.DefaultPubDirection
ms.assetid: 628352c1-040f-9ab1-d0f1-308b2c26679c
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.DefaultPubDirection property (Publisher)

Returns or sets a **[PbDirectionType](Publisher.PbDirectionType.md)** constant that represents the default direction in which text flows when a new publication is created. Read/write.


## Syntax

_expression_.**DefaultPubDirection**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

PbDirectionType


## Remarks

The **DefaultPubDirection** property value can be one of the **PbDirectionType** constants declared in the Microsoft Publisher type library.

This property generates an error if you are not running a bi-directional-enabled version of Microsoft Publisher (for example, Arabic).


## Example

This example sets the default direction for new publications and text flow in a bi-directional-enabled version of Publisher.

```vb
Sub SetDefaultDirection() 
 With Options 
 .DefaultPubDirection = pbDirectionRightToLeft 
 .DefaultTextFlowDirection = pbDirectionRightToLeft 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: Options.DefaultTextFlowDirection property (Publisher)
keywords: vbapb10.chm1048628
f1_keywords:
- vbapb10.chm1048628
ms.prod: publisher
api_name:
- Publisher.Options.DefaultTextFlowDirection
ms.assetid: 7c17768a-cd9c-704d-fa27-f0dfd7648054
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.DefaultTextFlowDirection property (Publisher)

Returns or sets a **[PbDirectionType](Publisher.PbDirectionType.md)** constant that represents a global Microsoft Publisher option, indicating whether text flows from left to right or from right to left in a publication. Read/write.


## Syntax

_expression_.**DefaultTextFlowDirection**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

PbDirectionType


## Remarks

The **DefaultTextFlowDirection** property value can be one of the **PbDirectionType** constants declared in the Publisher type library.

This property generates an error if you are not running a bi-directional-enabled version of Publisher (for example, Arabic).


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
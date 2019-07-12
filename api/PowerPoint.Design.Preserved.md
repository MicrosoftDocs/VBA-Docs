---
title: Design.Preserved property (PowerPoint)
keywords: vbapp10.chm644009
f1_keywords:
- vbapp10.chm644009
ms.prod: powerpoint
api_name:
- PowerPoint.Design.Preserved
ms.assetid: c7620e5a-49f5-49bc-307b-230ead112cf6
ms.date: 06/08/2017
localization_priority: Normal
---


# Design.Preserved property (PowerPoint)

Represents whether a design master is preserved from changes. Read/write.


## Syntax

_expression_. `Preserved`

_expression_ A variable that represents a [Design](PowerPoint.Design.md) object.


## Return value

MsoTriState


## Remarks

The value of the  **Preserved** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The design master is not preserved and can be edited.|
|**msoTrue**| The design master is preserved and cannot be edited.|

## Example

The following line of code locks and preserves the first design master.


```vb
Sub PreserveMaster

    ActivePresentation.Designs(1).Preserved = msoTrue

End Sub
```


## See also


[Design Object](PowerPoint.Design.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
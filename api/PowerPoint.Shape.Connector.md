---
title: Shape.Connector property (PowerPoint)
keywords: vbapp10.chm547020
f1_keywords:
- vbapp10.chm547020
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Connector
ms.assetid: 3e8cc3be-30a6-4e4e-32ca-bfd55ae973c2
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Connector property (PowerPoint)

Determines whether the specified shape is a connector. Read-only.


## Syntax

_expression_.**Connector**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

MsoTriState


## Remarks

The value of the  **Connector** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified shape is not a connector.|
|**msoTrue**| The specified shape is a connector.|

## Example

This example deletes all connectors on _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    For i = .Count To 1 Step -1

        With .Item(i)

            If .Connector Then .Delete

        End With

    Next

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
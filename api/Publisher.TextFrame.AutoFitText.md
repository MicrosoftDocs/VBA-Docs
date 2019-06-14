---
title: TextFrame.AutoFitText property (Publisher)
keywords: vbapb10.chm3866630
f1_keywords:
- vbapb10.chm3866630
ms.prod: publisher
api_name:
- Publisher.TextFrame.AutoFitText
ms.assetid: 468a9d3e-cb9d-8147-60ea-eb839d691e7a
ms.date: 06/15/2019
localization_priority: Normal
---


# TextFrame.AutoFitText property (Publisher)

Sets or returns a **[PbTextAutoFitType](Publisher.PbTextAutoFitType.md)** constant that represents how Microsoft Publisher automatically adjusts the text font size and the **TextFrame** object size for best viewing. Read/write.


## Syntax

_expression_.**AutoFitText**

_expression_ A variable that represents a **[TextFrame](Publisher.TextFrame.md)** object.


## Return value

PbTextAutoFitType


## Remarks

The **AutoFitText** property value can be one of the **PbTextAutoFitType** constants declared in the Microsoft Publisher type library.


## Example

The following example tests to see if the text frame has text, and if so, the **AutoFitText** property is set to best fit.

```vb
Sub TextFit() 
 
 Dim tfFrame As TextFrame 
 
 tfFrame = Application.ActiveDocument.MasterPages.Item(1).Shapes(1).TextFrame 
 With tfFrame 
 If .HasText = msoTrue Then .AutoFitText = pbTextAutoFitBestFit 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
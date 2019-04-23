---
title: NavigationControl.NumeralShapes property (Access)
keywords: vbaac10.chm11134
f1_keywords:
- vbaac10.chm11134
ms.prod: access
api_name:
- Access.NavigationControl.NumeralShapes
ms.assetid: 207bbece-366e-bc72-876f-98c80f7bf6b5
ms.date: 03/02/2019
localization_priority: Normal
---


# NavigationControl.NumeralShapes property (Access)

## Syntax

_expression_.**NumeralShapes**

_expression_ A variable that represents a **[NavigationControl](Access.NavigationControl.md)** object.


## Remarks

The **NumeralShapes** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|System|0|Numeral shapes are determined by the **Numeral Shapes** system setting.|
|Arabic|1|Arabic digit shapes are used to display and print numerals.|
|National|2|National digit shapes are used to display and print numerals.|
|Context|3|Numeral shapes are determined by Unicode context rules for adjacent text.|

## Example

The following example changes the **NumeralShapes** property for the selected control to 0 (numeral shapes will be determined by the **Numeral Shapes** system setting).

```vb
Public Sub ChangeNumeralShapes(ctl As Control) 
 ctl.NumeralShapes = 0 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
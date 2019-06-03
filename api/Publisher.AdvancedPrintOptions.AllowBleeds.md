---
title: AdvancedPrintOptions.AllowBleeds property (Publisher)
keywords: vbapb10.chm7077906
f1_keywords:
- vbapb10.chm7077906
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.AllowBleeds
ms.assetid: 0c12a611-4e1e-468b-ada2-f07d01fd4445
ms.date: 06/04/2019
localization_priority: Normal
---


# AdvancedPrintOptions.AllowBleeds property (Publisher)

**True** to allow bleeds to print for the specified publication. The default is **True**. Read/write **Boolean**.


## Syntax

_expression_.**AllowBleeds**

_expression_ A variable that represents an **[AdvancedPrintOptions](Publisher.AdvancedPrintOptions.md)** object.


## Return value

Boolean


## Remarks

When bleeds are allowed, objects that are partially off the page print to one eighth inch outside the defined page size.

If you allow bleeds in a document, you can specify whether bleed marks are printed by using the **[PrintBleedMarks](Publisher.AdvancedPrintOptions.PrintBleedMarks.md)** property.

This property corresponds to the **Allow bleeds** control on the **Page Settings** tab of the **Advanced Print Settings** dialog box.


## Example

The following example sets the publication to allow bleeds, and to print bleed marks.

```vb
Sub AllowBleedsAndPrintMarks() 
 With ActiveDocument.AdvancedPrintOptions 
 .AllowBleeds = True 
 .PrintBleedMarks = True 
 End With 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
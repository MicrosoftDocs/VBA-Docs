---
title: AdvancedPrintOptions.PrintBleedMarks property (Publisher)
keywords: vbapb10.chm7077907
f1_keywords:
- vbapb10.chm7077907
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.PrintBleedMarks
ms.assetid: f0c69d5f-4bfd-7a4c-3607-714859bcc86c
ms.date: 06/04/2019
localization_priority: Normal
---


# AdvancedPrintOptions.PrintBleedMarks property (Publisher)

**True** to print bleed marks in the specified publication. The default is **False**. Read/write **Boolean**.


## Syntax

_expression_.**PrintBleedMarks**

_expression_ A variable that represents an **[AdvancedPrintOptions](Publisher.AdvancedPrintOptions.md)** object.


## Return value

Boolean


## Remarks

Bleed marks show the extent of a bleed, and print an eighth inch outside the crop marks.

This property is only accessible if bleeds are allowed in the specified publication. Use the **[AllowBleeds](Publisher.AdvancedPrintOptions.AllowBleeds.md)** property to specify that bleeds are allowed. Returns "Permission Denied" if bleeds are not allowed in the publication.

This property corresponds to the **Bleed marks** control on the **Page Settings** tab of the **Advanced Print Settings** dialog box.


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
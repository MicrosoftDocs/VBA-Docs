---
title: InvisibleApp.DialogFont property (Visio)
keywords: vis_sdr.chm17552075
f1_keywords:
- vis_sdr.chm17552075
ms.prod: visio
api_name:
- Visio.InvisibleApp.DialogFont
ms.assetid: b9784c9b-99a5-7a48-01eb-dafbe6b2c4f9
ms.date: 06/25/2019
localization_priority: Normal
---


# InvisibleApp.DialogFont property (Visio)

Returns information about the fonts that Microsoft Visio uses in its dialog boxes. Read-only.


## Syntax

_expression_.**DialogFont**

_expression_ A variable that represents an **[InvisibleApp](Visio.InvisibleApp.md)** object.


## Return value

IFontDisp


## Remarks

You can use this property to display your dialog boxes in the same font as the Visio dialog boxes.

COM (Component Object Model) provides a standard implementation of a font object with the **IFontDisp** interface on top of the underlying system font support. The **IFontDisp** interface exposes a font object's properties and is implemented in the stdole type library as an **StdFont** object that can be created within Microsoft Visual Basic. The stdole type library is automatically referenced from all Visual Basic projects in Visio.

**To get information about the StdFont object that supports the IFontDisp interface**

1. In the **Code** group on the [Developer](../visio/How-to/run-visio-in-developer-mode.md) tab, choose **Visual Basic**.
    
2. On the **View** menu, choose **Object Browser**.
    
3. In the **Project/Library** list, choose **stdole**.
    
4. Under **Classes**, examine the class named **StdFont**.
    

## Example

The following sample code shows how to get a reference to an **StdFont** object that conveys information about the application fonts, and how to print that information to the Immediate window.

```vb
 
Sub DialogFont_Example() 
 
Dim objStdFont As StdFont 
Set objStdFont = Application.DialogFont 
 
 With objStdFont 
 
 Debug.Print .Bold 
 Debug.Print .CharSet 
 Debug.Print .Italic 
 Debug.Print .Name 
 Debug.Print .Size 
 Debug.Print .Strikethrough 
 Debug.Print .Underline 
 Debug.Print .Weight 
 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
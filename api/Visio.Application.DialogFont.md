---
title: Application.DialogFont Property (Visio)
keywords: vis_sdr.chm10052075
f1_keywords:
- vis_sdr.chm10052075
ms.prod: visio
api_name:
- Visio.Application.DialogFont
ms.assetid: 8742b97f-7f66-38c7-fafd-a343c1160671
ms.date: 06/08/2017
---


# Application.DialogFont Property (Visio)

Returns information about the fonts that Microsoft Visio uses in its dialog boxes. Read-only.


## Syntax

 _expression_. `DialogFont`

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


### Return value

IFontDisp


## Remarks

You can use this property to display your dialog boxes in the same font as the Visio dialog boxes.

COM (Component Object Model) provides a standard implementation of a font object with the  **IFontDisp** interface on top of the underlying system font support. The **IFontDisp** interface exposes a font object's properties and is implemented in the stdole type library as a **StdFont** object that can be created within Microsoft Visual Basic. The stdole type library is automatically referenced from all Visual Basic projects in Visio.

 **To get information about the StdFont object that supports the IFontDisp interface**




1. In the  **Code** group on the [Developer](../visio/How-to/run-visio-in-developer-mode.md) tab, click **Visual Basic**.
    
2. On the  **View** menu, click **Object Browser**.
    
3. In the  **Project/Library** list, click **stdole**.
    
4. Under  **Classes**, examine the class named  **StdFont** .
    


For details about the  **IFontDisp** interface, see the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.


## Example

The following sample code shows how to get a reference to a  **StdFont** object that conveys information about the application fonts, and how to print that information to the Immediate window.


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



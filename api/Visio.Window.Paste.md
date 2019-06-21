---
title: Window.Paste method (Visio)
ms.prod: visio
api_name:
- Visio.Window.Paste
ms.assetid: e5535c75-5a43-48dc-bd77-50db003809ba
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.Paste method (Visio)

This object, member, or enumeration is deprecated and is not intended to be used in your code. Pastes the contents of the Clipboard into an object.


## Version Information

Version Added: Visio 2.0 


### Syntax

_expression_.**Paste**

_expression_ A variable that represents a **[Window](Visio.Window.md)** object.


## Remarks

The  **Window** object's **Paste** method is now obsolete. Use the **Paste** or **PasteSpecial** method of the [Page](Visio.Page.md), [Master](Visio.Master.md), or [Shape](Visio.Shape.md) object. (Use the **Shape** object in the case of group shapes.)

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVWindow.Paste()**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
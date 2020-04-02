---
title: Global object (Visio)
keywords: vis_sdr.chm10000
f1_keywords:
- vis_sdr.chm10000
ms.prod: visio
api_name:
- Visio.Global
ms.assetid: 3c7dca10-f7b0-f3f7-59b1-7845338aa4a4
ms.date: 06/19/2019
localization_priority: Normal
---


# Global object (Visio)

The Microsoft Visio **Global** object is automatically available to Microsoft Visual Basic for Applications (VBA) code that is part of the VBA project of a Visio document. The **Global** object is not available to code in other contexts.


## Remarks

Members of the **Global** object can be accessed without qualification. For example, to access the **ActivePage** member of the **Global** object, use the following code.

```vb
    Set vsoPage = ActivePage 
```

The preceding syntax is different from the syntax that you would use for accessing members of non-global objects. For example:

```vb
    Set vsoPage = vsoApplication.ActivePage
```

> [!NOTE] 
> The VBA project of every Visio document also has a class module called **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)**. When referenced from code in the VBA project, the **ThisDocument** module returns a reference to the project's **Document** object.

## Properties

- [ActiveDocument](Visio.Global.ActiveDocument.md)
- [ActivePage](Visio.Global.ActivePage.md)
- [ActiveWindow](Visio.Global.ActiveWindow.md)
- [Addons](Visio.Global.Addons.md)
- [Application](Visio.Global.Application.md)
- [Documents](Visio.Global.Documents.md)
- [VBE](Visio.Global.Vbe.md)
- [Windows](Visio.Global.Windows.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
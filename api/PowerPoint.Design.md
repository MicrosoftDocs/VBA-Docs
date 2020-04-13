---
title: Design object (PowerPoint)
keywords: vbapp10.chm644000
f1_keywords:
- vbapp10.chm644000
ms.prod: powerpoint
api_name:
- PowerPoint.Design
ms.assetid: 3b02c779-8313-9512-c8d9-cf8a3883229f
ms.date: 06/08/2017
localization_priority: Normal
---


# Design object (PowerPoint)

Represents an individual slide design template. The **Design** object is a member of the **[Designs](PowerPoint.Designs.md)** and **[SlideRange](PowerPoint.SlideRange.md)** collections and the **[Master](PowerPoint.Master.md)** and **[Slide](PowerPoint.Slide.md)** objects.


## Remarks

Use the  **Design** property of the **Master**, **Slide**, or **SlideRange** objects to access a **Design** object, for example:


-  `ActivePresentation.SlideMaster.Design`
    
-  `ActivePresentation.Slides(1).Design`
    
-  `ActivePresentation.Slides.Range.Design`
    
Use the [Add](PowerPoint.Designs.Add.md), [Item](PowerPoint.Designs.Item.md), [Clone](PowerPoint.Designs.Clone.md), or [Load](PowerPoint.Designs.Load.md)methods of the  **Designs** collection to add, refer to, clone, or load a **Design** object, respectively. For example, to add a design template, use `ActivePresentation.Designs.Add designName:="MyDesign"`


## Example

The **Design** object's[AddTitleMaster](PowerPoint.Presentation.AddTitleMaster.md)method and [HasTitleMaster](PowerPoint.Presentation.HasTitleMaster.md)property can be used to add and / or query the status of a title slide master. For example:


```vb
Sub AddQueryTitleMaster(dsn As Design)

    dsn.AddTitleMaster

    MsgBox dsn.HasTitleMaster

End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
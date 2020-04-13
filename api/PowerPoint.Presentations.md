---
title: Presentations object (PowerPoint)
keywords: vbapp10.chm522000
f1_keywords:
- vbapp10.chm522000
ms.prod: powerpoint
api_name:
- PowerPoint.Presentations
ms.assetid: 0b952edc-8628-71ef-e854-3bcefbb3bc61
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentations object (PowerPoint)

A collection of all the  **[Presentation](PowerPoint.Presentation.md)** objects in Microsoft PowerPoint. Each **Presentation** object represents a presentation that's currently open in PowerPoint.


## Remarks

The **Presentations** collection doesn't include open add-ins, which are a special kind of hidden presentation. You can, however, return a single open add-in if you know its file name. For example `Presentations("oscar.ppa")` will return the open add-in named "Oscar.ppa" as a **Presentation** object. However, it is recommended that the **AddIns** collection be used to return open add-ins.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint.Presentations.GetEnumerator** (to enumerate the **Presentation** objects.)
    

## Example

Use the [Presentations](PowerPoint.Application.Presentations.md) property to return the **Presentations** collection. Use the [Add](PowerPoint.Presentations.Add.md) method to create a new presentation and add it to the collection. The following example creates a new presentation, adds a slide to the presentation, and then saves the presentation.


```vb
Set newPres = Presentations.Add(True) 
newPres.Slides.Add 1, 1 
newPres.SaveAs "Sample"
```

Use  **Presentations** (_index_), where _index_ is the presentation's name or index number, to return a single **Presentation** object. The following example prints presentation one.




```vb
Presentations(1).PrintOut
```

Use the [Open](PowerPoint.Presentations.Open.md) method to open a presentation and add it to the **Presentations** collection. The following example opens the file Sales.ppt as a read-only presentation.




```vb
Presentations.Open FileName:="sales.ppt", ReadOnly:=True
```


## Methods



|Name|
|:-----|
|[Add](PowerPoint.Presentations.Add.md)|
|[CanCheckOut](PowerPoint.Presentations.CanCheckOut.md)|
|[CheckOut](PowerPoint.Presentations.CheckOut.md)|
|[Item](PowerPoint.Presentations.Item.md)|
|[Open](PowerPoint.Presentations.Open.md)|
|[Open2007](PowerPoint.Presentations.Open2007.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.Presentations.Application.md)|
|[Count](PowerPoint.Presentations.Count.md)|
|[Parent](PowerPoint.Presentations.Parent.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

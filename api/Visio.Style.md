---
title: Style object (Visio)
keywords: vis_sdr.chm10255
f1_keywords:
- vis_sdr.chm10255
ms.prod: visio
api_name:
- Visio.Style
ms.assetid: fdebb8d1-8910-3df8-74cd-9f847efb7ecb
ms.date: 06/19/2019
localization_priority: Normal
---


# Style object (Visio)

Represents a style defined in a document.


## Remarks

You retrieve a particular style from the **[Styles](Visio.Styles.md)** collection of a **Document** object.

The default property of a **Style** object is **Name**.

Any **[Shape](visio.shape.md)** object to which a style is applied inherits the attributes defined by the style. Use the **LineStyle**, **FillStyle**, **TextStyle**, or **Style** property of a **Shape** object to apply a style to a shape or to determine what style is applied to a shape.

Like a **Shape** object, a **Style** object has cells whose formulas define the values of the style's attributes. To retrieve one of these cells, use the **Cells** or **CellsSRC** property.

## Events

-  [BeforeStyleDelete](Visio.Style.BeforeStyleDelete.md)
-  [QueryCancelStyleDelete](Visio.Style.QueryCancelStyleDelete.md)
-  [StyleChanged](Visio.Style.StyleChanged.md)
-  [StyleDeleteCanceled](Visio.Style.StyleDeleteCanceled.md)

## Methods

-  [Delete](Visio.Style.Delete.md)
-  [GetFormulas](Visio.Style.GetFormulas.md)
-  [GetFormulasU](Visio.Style.GetFormulasU.md)
-  [GetResults](Visio.Style.GetResults.md)
-  [SetFormulas](Visio.Style.SetFormulas.md)
-  [SetResults](Visio.Style.SetResults.md)

## Properties

-  [Application](Visio.Style.Application.md)
-  [BasedOn](Visio.Style.BasedOn.md)
-  [CellExists](Visio.Style.CellExists.md)
-  [CellExistsU](Visio.Style.CellExistsU.md)
-  [Cells](Visio.Style.Cells.md)
-  [CellsSRC](Visio.Style.CellsSRC.md)
-  [CellsSRCExists](Visio.Style.CellsSRCExists.md)
-  [CellsU](Visio.Style.CellsU.md)
-  [Document](Visio.Style.Document.md)
-  [EventList](Visio.Style.EventList.md)
-  [FillBasedOn](Visio.Style.FillBasedOn.md)
-  [Hidden](Visio.Style.Hidden.md)
-  [ID](Visio.Style.ID.md)
-  [IncludesFill](Visio.Style.IncludesFill.md)
-  [IncludesLine](Visio.Style.IncludesLine.md)
-  [IncludesText](Visio.Style.IncludesText.md)
-  [Index](Visio.Style.Index.md)
-  [LineBasedOn](Visio.Style.LineBasedOn.md)
-  [Name](Visio.Style.Name.md)
-  [NameU](Visio.Style.NameU.md)
-  [ObjectType](Visio.Style.ObjectType.md)
-  [PersistsEvents](Visio.Style.PersistsEvents.md)
-  [Section](Visio.Style.Section.md)
-  [Stat](Visio.Style.Stat.md)
-  [TextBasedOn](Visio.Style.TextBasedOn.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
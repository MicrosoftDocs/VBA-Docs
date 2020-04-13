---
title: Connect object (Visio)
keywords: vis_sdr.chm10065
f1_keywords:
- vis_sdr.chm10065
ms.prod: visio
api_name:
- Visio.Connect
ms.assetid: f29481d6-ceaa-69b4-5e44-26e06199488d
ms.date: 06/19/2019
localization_priority: Normal
---


# Connect object (Visio)

Represents a connection between two shapes in a drawing, such as a line and a box in an organization chart.


## Remarks

Retrieve a **Connect** object from the **[Connects](Visio.Connects.md)** collection returned by the **Connects** and **FromConnects** properties of a **[Shape](visio.shape.md)** object, or from the **[Page.Connects](visio.page.connects.md)** or **[Master.Connects](visio.master.connects.md)** property.

The default property of a **Connect** object is **FromSheet**.

Use the **GlueTo** or **GlueToPos** method of a **[Cell](visio.cell.md)** object to connect one shape to another in a drawing.


## Properties

- [Application](Visio.Connect.Application.md)
- [ContainingMasterID](Visio.Connect.ContainingMasterID.md)
- [ContainingPageID](Visio.Connect.ContainingPageID.md)
- [Document](Visio.Connect.Document.md)
- [FromCell](Visio.Connect.FromCell.md)
- [FromPart](Visio.Connect.FromPart.md)
- [FromSheet](Visio.Connect.FromSheet.md)
- [Index](Visio.Connect.Index.md)
- [ObjectType](Visio.Connect.ObjectType.md)
- [Stat](Visio.Connect.Stat.md)
- [ToCell](Visio.Connect.ToCell.md)
- [ToPart](Visio.Connect.ToPart.md)
- [ToSheet](Visio.Connect.ToSheet.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
---
title: ContainerProperties object (Visio)
keywords: vis_sdr.chm61040
f1_keywords:
- vis_sdr.chm61040
ms.prod: visio
api_name:
- Visio.ContainerProperties
ms.assetid: b94f758f-58f7-f1ef-c03b-761e26c11017
ms.date: 06/19/2019
localization_priority: Normal
---


# ContainerProperties object (Visio)

Represents the set of properties that are specific to a container.


## Remarks

A container is a shape that has some specific properties that are represented conceptually by a **ContainerProperties** object. Containers appear as collections of shapes surrounded by borders that are sometimes, but not always, visible. Some, but not all, containers are lists.

You can use the **[Shape.ContainerProperties](Visio.Shape.ContainerProperties.md)** property to get a **ContainerProperties** object that represents the set of container properties of any shape, whether or not it is actually a container.

To create an instance of a container shape on the drawing page, use the **[Page.DropContainer](Visio.Page.DropContainer.md)** method.


## Methods

- [AddMember](Visio.ContainerProperties.AddMember.md)
- [Disband](Visio.ContainerProperties.Disband.md)
- [FitToContents](Visio.ContainerProperties.FitToContents.md)
- [GetListMemberPosition](Visio.ContainerProperties.GetListMemberPosition.md)
- [GetListMembers](Visio.ContainerProperties.GetListMembers.md)
- [GetListSpacing](Visio.ContainerProperties.GetListSpacing.md)
- [GetMargin](Visio.ContainerProperties.GetMargin.md)
- [GetMemberShapes](Visio.ContainerProperties.GetMemberShapes.md)
- [GetMemberState](Visio.ContainerProperties.GetMemberState.md)
- [InsertListMember](Visio.ContainerProperties.InsertListMember.md)
- [RemoveMember](Visio.ContainerProperties.RemoveMember.md)
- [ReorderListMember](Visio.ContainerProperties.ReorderListMember.md)
- [RotateFlipList](Visio.ContainerProperties.RotateFlipList.md)
- [SetListSpacing](Visio.ContainerProperties.SetListSpacing.md)
- [SetMargin](Visio.ContainerProperties.SetMargin.md)


## Properties

- [Application](Visio.ContainerProperties.Application.md)
- [ContainerStyle](Visio.ContainerProperties.ContainerStyle.md)
- [ContainerType](Visio.ContainerProperties.ContainerType.md)
- [Document](Visio.containerproperties.document.md)
- [HeadingStyle](Visio.ContainerProperties.HeadingStyle.md)
- [ListAlignment](Visio.ContainerProperties.ListAlignment.md)
- [ListDirection](Visio.ContainerProperties.ListDirection.md)
- [LockMembership](Visio.ContainerProperties.LockMembership.md)
- [ObjectType](Visio.ContainerProperties.ObjectType.md)
- [OverlappedList](Visio.ContainerProperties.OverlappedList.md)
- [ResizeAsNeeded](Visio.ContainerProperties.ResizeAsNeeded.md)
- [Shape](Visio.ContainerProperties.Shape.md)
- [Stat](Visio.ContainerProperties.Stat.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
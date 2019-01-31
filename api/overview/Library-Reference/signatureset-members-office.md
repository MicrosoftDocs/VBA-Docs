---
title: SignatureSet members (Office)
description: A collection of Signature objects that correspond to the digital signature attached to a document.
ms.prod: office
ms.assetid: abe810a3-ffe4-ee26-8df7-d68cfbf3bf1e
ms.date: 01/30/2019
localization_priority: Normal
---


# SignatureSet members (Office)

A collection of **Signature** objects that correspond to the digital signature attached to a document.


## Methods

|Name|Description|
|:-----|:-----|
|[AddNonVisibleSignature](../../Office.SignatureSet.AddNonVisibleSignature.md)|Creates a signature packet when digitally signing a document.|
|[AddSignatureLine](../../Office.SignatureSet.AddSignatureLine.md)|Adds lines to a document where signatures are collected.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.SignatureSet.Application.md)|Gets an **Application** object that represents the container application for the **SignatureSet** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[CanAddSignatureLine](../../Office.SignatureSet.CanAddSignatureLine.md)|Gets a **Boolean** value indicating whether you can add a signature line to a document. Read-only.|
|[Count](../../Office.SignatureSet.Count.md)|Gets a **Long** indicating the number of items in the **SignatureSet** object. Read-only.|
|[Creator](../../Office.SignatureSet.Creator.md)|Gets a 32-bit integer that indicates the application in which the **SignatureSet** object was created. Read-only.|
|[Item](../../Office.SignatureSet.Item.md)|Gets a **Signature** object that corresponds to one of the digital signatures with which the document is currently signed. Read-only.|
|[Parent](../../Office.SignatureSet.Parent.md)|Gets the **Parent** object for the **SignatureSet** object. Read-only.|
|[ShowSignaturesPane](../../Office.SignatureSet.ShowSignaturesPane.md)|Gets or sets a **Boolean** value indicating whether the **Signature** task pane should be displayed. Read/write.|
|[Subset](../../Office.SignatureSet.Subset.md)|Gets or sets a value that acts as a filter on the available **Signature** objects for a document. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
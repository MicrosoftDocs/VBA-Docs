---
title: Signature members (Office)
description: Represents a digital signature attached to a document. Signature objects are contained in the SignatureSet collection of the Document object.
ms.prod: office
ms.assetid: 1054db23-fe1c-f81f-e44b-d8c2c82ca7fa
ms.date: 01/30/2019
localization_priority: Normal
---


# Signature members (Office)

Represents a digital signature attached to a document. **Signature** objects are contained in the **SignatureSet** collection of the **Document** object.


## Methods

|Name|Description|
|:-----|:-----|
|[Delete](../../Office.Signature.Delete.md)|Deletes the **Signature** object from the collection.|
|[ShowDetails](../../Office.Signature.ShowDetails.md)|Displays details related to a signature packet.|
|[Sign](../../Office.Signature.Sign.md)|Creates a signature packet.|

## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.Signature.Application.md)|Gets an **Application** object that represents the container application for the **Signature** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[CanSetup](../../Office.Signature.CanSetup.md)|Gets a **Boolean** value indicating whether the user can set properties of the **Signature** object. Read-only.|
|[Creator](../../Office.Signature.Creator.md)|Gets a 32-bit integer that indicates the application in which the **Signature** object was created. Read-only.|
|[Details](../../Office.Signature.Details.md)|Gets information about a signature. Read-only.|
|[IsSignatureLine](../../Office.Signature.IsSignatureLine.md)|Gets a value indicating whether this is a signature line. Read-only.|
|[IsSigned](../../Office.Signature.IsSigned.md)|Gets a **Boolean** value indicating whether the document was signed successfully. Read-only.|
|[Parent](../../Office.Signature.Parent.md)|Gets the **Parent** object for the Signature object. Read-only.|
|[Setup](../../Office.Signature.Setup.md)|Gets a **SignatureSetup** object that provides access to various properties of a signature packet. Read-only.|
|[SignatureLineShape](../../Office.Signature.SignatureLineShape.md)|Gets the **Shape** object associated with a **Signature** object that is a signature line. Read-only.|
|[SortHint](../../Office.Signature.SortHint.md)|Gets a value representing the sort order of the signatures in a packet with multiple signatures. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
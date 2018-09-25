---
title: Windows Collection (VBA Add-In Object Model)
keywords: vbob6.chm1071203
f1_keywords:
- vbob6.chm1071203
ms.prod: office
ms.assetid: 5f758e82-f571-e62d-2d35-c0917cbe0f59
ms.date: 06/08/2017
---


# Windows Collection (VBA Add-In Object Model)



Contains all open or permanent windows.

## Remarks

Use the  **Windows** collection to access **Window** objects.
The  **Windows** collection has a fixed set of windows that are always available in the [collection](../../Glossary/vbe-glossary.md#collection), such as the [Project window](../../Glossary/vbe-glossary.md#project-window), the [Properties window](../../Glossary/vbe-glossary.md#properties-window), and a set of windows that represent all open code windows and [designer](../../Glossary/vbe-glossary.md#designer) windows. Opening a code or designer window adds a new member to the **Windows** collection. Closing a code or designer window removes a member from the **Windows** collection. Closing a permanent[development environment](../../Glossary/vbe-glossary.md#development-environment) window doesn't remove the corresponding object from this collection, but results in the window not being visible.


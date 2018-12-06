---
title: CodeModule object (VBA Add-In Object Model)
keywords: vbob6.chm1070944
f1_keywords:
- vbob6.chm1070944
ms.prod: office
ms.assetid: f2ce876d-ee2b-058f-37fc-f681bd41f139
ms.date: 12/06/2018
---


# CodeModule object (VBA Add-In Object Model)

Represents the code behind a component, such as a [form](../../Glossary/vbe-glossary.md#form), [class](../../Glossary/vbe-glossary.md#class), or [document](../../Glossary/vbe-glossary.md#document).

## Remarks

You use the **CodeModule** object to modify (add, delete, or edit) the code associated with a component. Each component is associated with one **CodeModule** object. However, a **CodeModule** object can be associated with multiple [code panes](../../Glossary/vbe-glossary.md#code-pane).

The methods associated with the **CodeModule** object enable you to manipulate and return information about the code text on a line-by-line basis. For example, you can use the **[AddFromString](addfromstring-method-vba-add-in-object-model.md)** method to add text to the [module](../../Glossary/vbe-glossary.md#module). **AddFromString** places the text just above the first [procedure](../../Glossary/vbe-glossary.md#procedure) in the module or places the text at the end of the module if there are no procedures.

Use the **[Parent](parent-property-vba-add-in-object-model.md)** property to return the **[VBComponent](vbcomponent-object-vba-add-in-object-model.md)** object associated with a [code module](../../Glossary/vbe-glossary.md#code-module).

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)
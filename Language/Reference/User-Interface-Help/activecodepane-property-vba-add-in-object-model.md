---
title: ActiveCodePane property (VBA Add-In Object Model)
keywords: vbob6.chm1070961
f1_keywords:
- vbob6.chm1070961
ms.prod: office
ms.assetid: 7c9839e2-e458-1dc5-f402-b05305503824
ms.date: 12/06/2018
---


# ActiveCodePane property (VBA Add-In Object Model)

Returns the active or last active **[CodePane](codepane-object-vba-add-in-object-model.md)** object or sets the active **CodePane** object. Read/write.

## Remarks

You can set the **ActiveCodePane** property to any valid **CodePane** object, as shown in the following example:

```vb
Set MyApp.VBE. ActiveCodePane = MyApp.VBE.CodePanes(1)

```

The preceding example sets the first [code pane](../../Glossary/vbe-glossary.md#code-pane) in a [collection](../../Glossary/vbe-glossary.md#collection) of code panes to be the active code pane. You can also activate a code pane by using the **Set** method.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)
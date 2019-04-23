---
title: Making faster For...Next loops (VBA)
keywords: vbcn6.chm1009794
f1_keywords:
- vbcn6.chm1009794
ms.prod: office
ms.assetid: 4a483362-fd6b-f0a7-5cb0-b85a2f794937
ms.date: 12/21/2018
localization_priority: Normal
---


# Making faster For...Next loops

Integers use less memory than the [Variant data type](../../Glossary/vbe-glossary.md#variant-data-type) and are slightly faster to update. However, this difference is only noticeable if you perform many thousands of operations. For example:

```vb
Dim CountFaster As Integer    ' First case, use Integer. 
For CountFaster = 0 to 32766     
Next CountFaster 
 
Dim CountSlower As Variant    ' Second case, use Variant. 
For CountSlower = 0 to 32766 
Next CountSlower 

```


The first case takes slightly less time to run than the second case. However, if `CountFaster` exceeds 32,767, an error occurs. To fix this, you can change `CountFaster` to the [Long data type](../../Glossary/vbe-glossary.md#long-data-type), which accepts a wider range of integers. In general, the smaller the [data type](../../reference/user-interface-help/data-type-summary.md), the less time it takes to update. Variants are slightly slower than their equivalent data type.

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]

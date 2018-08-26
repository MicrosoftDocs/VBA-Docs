---
title: Me <keyword>
keywords: vblr6.chm1008868
f1_keywords:
- vblr6.chm1008868
ms.prod: office
ms.assetid: 6d062019-bb49-7acb-5f03-7bb5a2a09681
ms.date: 06/08/2017
---


# Me <keyword>

<<<<<<< HEAD
The  **Me**[keyword](../../Glossary/vbe-glossary.md) behaves like an implicitly declared[variable](../../Glossary/vbe-glossary.md). It is automatically available to every [procedure](../../Glossary/vbe-glossary.md) in a[class module](../../Glossary/vbe-glossary.md). When a [class](../../Glossary/vbe-glossary.md) can have more than one instance, **Me** provides a way to refer to the specific instance of the class where the code is executing. Using **Me** is particularly useful for passing information about the currently executing instance of a class to a procedure in another[module](../../Glossary/vbe-glossary.md). For example, suppose you have the following procedure in a module:
=======
The  **Me**[keyword](../../Glossary/vbe-glossary.md#keyword) behaves like an implicitly declared[variable](../../Glossary/vbe-glossary.md#variable). It is automatically available to every [procedure](../../Glossary/vbe-glossary.md#procedure) in a[class module](../../Glossary/vbe-glossary.md#class-module). When a [class](../../Glossary/vbe-glossary.md#class) can have more than one instance, **Me** provides a way to refer to the specific instance of the class where the code is executing. Using **Me** is particularly useful for passing information about the currently executing instance of a class to a procedure in another[module](../../Glossary/vbe-glossary.md#module). For example, suppose you have the following procedure in a module:
>>>>>>> master


```vb
Sub ChangeFormColor(FormName As Form) 
 FormName.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256) 
End Sub
```


<<<<<<< HEAD
You can call this procedure and pass the current instance of the Form class as an [argument](../../Glossary/vbe-glossary.md) using the following[statement](../../Glossary/vbe-glossary.md):
=======
You can call this procedure and pass the current instance of the Form class as an [argument](../../Glossary/vbe-glossary.md#argument) using the following[statement](../../Glossary/vbe-glossary.md#statement):
>>>>>>> master




```vb
ChangeFormColor Me 

```



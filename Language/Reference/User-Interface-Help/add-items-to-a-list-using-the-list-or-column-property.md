---
title: Add items to a list using the List or Column property
keywords: fm20.chm5225253
f1_keywords:
- fm20.chm5225253
ms.prod: office
ms.assetid: 08757f51-9c54-9ef7-7268-48824515b020
ms.date: 06/08/2017
---


# Add items to a list using the List or Column property




1. Create a multicolumn  **[ListBox](listbox-control.md)** or **[ComboBox](combobox-control.md)**.
    
2. Create a two-dimensional [array](../../Glossary/vbe-glossary.md#array) that contains the items you want to put in the list.
    
3. Set the  **ColumnCount** property of the **[ListBox](listbox-control.md)** or **[ComboBox](combobox-control.md)** to match the number of entries in the list.
    
4. Do one of the following:
    
    
    
      - Assign the array as the value of the  **List** property. The contents of the **[ListBox](listbox-control.md)** will match the contents of the array exactly.
    
  - Assign the array as the value of the  **Column** property. **Column** transposes rows and columns, so each row of the **[ListBox](listbox-control.md)** matches the corresponding column of the array.
    

    
    





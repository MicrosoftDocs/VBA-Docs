---
title: Tips on using text boxes
keywords: fm20.chm5225199
f1_keywords:
- fm20.chm5225199
ms.prod: office
ms.assetid: 4f2c565a-50ba-0295-d8bf-92d316ea25af
ms.date: 12/29/2018
localization_priority: Normal
---


# Tips on using text boxes

The **[TextBox](../../reference/user-interface-help/textbox-control.md)** is a flexible control governed by the following properties: **Text**, **MultiLine**, **WordWrap**, and **AutoSize**.

- **[Text](../../reference/user-interface-help/text-property-microsoft-forms.md)** contains the text that's displayed in the text box.

- **[MultiLine](../../reference/user-interface-help/multiline-property.md)** controls whether the **TextBox** can display text as a single line or as multiple lines. Newline characters identify where one line ends and another begins. If **MultiLine** is **False**, the text is truncated instead of wrapped.

- **[WordWrap](../../reference/user-interface-help/wordwrap-property.md)** allows the **TextBox** to wrap lines of text that are longer than the width of the **TextBox** into shorter lines that fit.

   If you do not use **WordWrap**, the **TextBox** starts a new line of text when it encounters a newline character in the text. If **WordWrap** is turned off, you can have text lines that do not fit completely in the **TextBox**. The **TextBox** displays the portions of text that fit inside its width and truncates the portions of text that do not fit. **WordWrap** is not applicable unless **MultiLine** is **True**.

- **[AutoSize](../../reference/user-interface-help/autosize-property.md)** controls whether the **TextBox** adjusts to display all of the text. When using **AutoSize** with a **TextBox**, the width of the **TextBox** shrinks or expands according to the amount of text in the **TextBox** and the font size used to display the text.
  
  **AutoSize** works well in the following situations:

    - Displaying a caption of one or more lines.   
    - Displaying the contents of a single-line **TextBox**.
    - Displaying the contents of a multiline **TextBox** that is read-only to the user.
    
> [!NOTE] 
> Avoid using **AutoSize** with an empty **TextBox** that also uses the **MultiLine** and **WordWrap** properties. When the user enters text into a **TextBox** with these properties, the **TextBox** automatically sizes to a long narrow box one character wide and as long as the line of text.

## See also

- [Microsoft Forms reference](../../reference/user-interface-help/reference-microsoft-forms.md)
- [Microsoft Forms conceptual topics](../../reference/user-interface-help/concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
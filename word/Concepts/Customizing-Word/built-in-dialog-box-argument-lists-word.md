---
title: Built-in Dialog Box Argument Lists (Word)
keywords: vbawd10.chm5210109
f1_keywords:
- vbawd10.chm5210109
ms.prod: word
ms.assetid: 0d8c9f48-85cc-6744-a699-668040313862
ms.date: 06/08/2017
localization_priority: Normal
---


# Built-in Dialog Box Argument Lists (Word)

Many of the built-in dialog boxes in Word have options that you may want to set. To set or return the properties associated with a Word dialog box, use the equivalent Visual Basic properties and methods. For example, if you want to print a document, use the VBA **[PrintOut](../../../api/Word.Document.PrintOut.md)** method. The following code prints the current document using the **Print** dialog box default settings. However, if you do not want to use the default setting in the print dialog, you can use the arguments associated with the **PrintOut** method.


```vb
Sub PrintCurrentDocument() 
 ActiveDocument.PrintOut 
End Sub
```


Although you are encouraged to use VBA keywords to get or set the value of dialog box options, many of the built-in Word dialog boxes have arguments that you can also use to get or set values from a dialog box. For more information, see  [Displaying Built-in Word Dialog Boxes](displaying-built-in-word-dialog-boxes.md).



|**WdWordDialog constant**|**Argument lists**|
|:-----|:-----|
| **wdDialogConnect**| **_Drive_** , **_Path_**, **_Password_**|
| **wdDialogConsistencyChecker**|(none)|
| **wdDialogControlRun**| **_Application_**|
| **wdDialogConvertObject**| **_IconNumber_** , **_ActivateAs_**, **_IconFileName_**, **_Caption_**, **_Class_**, **_DisplayIcon_**, **_Floating_**|
| **wdDialogCopyFile**| **_FileName_** , **_Directory_**|
| **wdDialogCreateAutoText**|(none)|
| **wdDialogCSSLinks**| **_LinkStyles_**|
| **wdDialogDocumentStatistics**| **_FileName_** , **_Directory_**, **_Template_**, **_Title_**, **_Created_**, **_LastSaved_**, **_LastSavedBy_**, **_Revision_**, **_Time_**, **_Printed_**, **_Pages_**, **_Words_**, **_Characters_**, **_Paragraphs_**, **_Lines_**, **_FileSize_**|
| **wdDialogDrawAlign**| **_Horizontal_** , **_Vertical_**, **_RelativeTo_**|
| **wdDialogDrawSnapToGrid**| **_SnapToGrid_** , **_XGrid_**, **_YGrid_**, **_XOrigin_**, **_YOrigin_**, **_SnapToShapes_**, **_XGridDisplay_**, **_YGridDisplay_**, **_FollowMargins_**, **_ViewGridLines_**, **_DefineLineBasedOnGrid_**|
| **wdDialogEditAutoText**| **_Name_** , **_Context_**, **_InsertAs_**, **_Insert_**, **_Add_**, **_Define_**, **_InsertAsText_**, **_Delete_**, **_CompleteAT_**|
| **wdDialogEditCreatePublisher**|(For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.)|
| **wdDialogEditFind**| **_Find_** , **_Replace_**, **_Direction_**, **_MatchCase_**, **_WholeWord_**, **_PatternMatch_**, **_SoundsLike_**, **_FindNext_**, **_ReplaceOne_**, **_ReplaceAll_**, **_Format_**, **_Wrap_**, **_FindAllWordForms_**, **_MatchByte_**, **_FuzzyFind_**, **_Destination_**, **_CorrectEnd_**, **_MatchKashida_**, **_MatchDiacritics_**, **_MatchAlefHamza_**, **_MatchControl_**|
| **wdDialogEditFrame**| **_Wrap_** , _WidthRule_, **_FixedWidth_**,  _HeightRule_, **_FixedHeight_**, **_PositionHorz_**, **_PositionHorzRel_**, **_DistFromText_**, **_PositionVert_**, **_PositionVertRel_**, **_DistVertFromText_**, **_MoveWithText_**,  _LockAnchor_, **_RemoveFrame_**|
| **wdDialogEditGoTo**| **_Find_** , **_Replace_**, **_Direction_**, **_MatchCase_**, **_WholeWord_**, **_PatternMatch_**, **_SoundsLike_**, **_FindNext_**, **_ReplaceOne_**, **_ReplaceAll_**, **_Format_**, **_Wrap_**, **_FindAllWordForms_**, **_MatchByte_**, **_FuzzyFind_**, **_Destination_**, **_CorrectEnd_**, **_MatchKashida_**, **_MatchDiacritics_**, **_MatchAlefHamza_**, **_MatchControl_**|
| **wdDialogEditGoToOld**|(none)|
| **wdDialogEditLinks**| **_UpdateMode_** , **_Locked_**, **_SavePictureInDoc_**, **_UpdateNow_**, **_OpenSource_**, **_KillLink_**, **_Link_**, **_Application_**, **_Item_**, **_FileName_**, **_PreserveFormatLinkUpdate_**|
| **wdDialogEditObject**| **_Verb_**|
| **wdDialogEditPasteSpecial**| **_IconNumber_** , **_Link_**, **_DisplayIcon_**, **_Class_**, **_DataType_**, **_IconFileName_**, **_Caption_**, **_Floating_**|
| **wdDialogEditPublishOptions**|(For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.)|
| **wdDialogEditReplace**| **_Find_** , **_Replace_**, **_Direction_**, **_MatchCase_**, **_WholeWord_**, **_PatternMatch_**, **_SoundsLike_**, **_FindNext_**, **_ReplaceOne_**, **_ReplaceAll_**, **_Format_**, **_Wrap_**, **_FindAllWordForms_**, **_MatchByte_**, **_FuzzyFind_**, **_Destination_**, **_CorrectEnd_**, **_MatchKashida_**, **_MatchDiacritics_**, **_MatchAlefHamza_**, **_MatchControl_**|
| **wdDialogEditStyle**|(none)|
| **wdDialogEditSubscribeOptions**|(For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.)|
| **wdDialogEditSubscribeTo**|(For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.)|
| **wdDialogEditTOACategory**| **_Category_** , **_CategoryName_**|
| **wdDialogEmailOptions**|(none)|
| **wdDialogFileDocumentLayout**| **_Tab_** , **_PaperSize_**, **_TopMargin_**, **_BottomMargin_**, **_LeftMargin_**, **_RightMargin_**, **_Gutter_**, **_PageWidth_**, **_PageHeight_**, **_Orientation_**, **_FirstPage_**, **_OtherPages_**, **_VertAlign_**, **_ApplyPropsTo_**, **_Default_**, **_FacingPages_**, **_HeaderDistance_**, **_FooterDistance_**, **_SectionStart_**, **_OddAndEvenPages_**, **_DifferentFirstPage_**, **_Endnotes_**, **_LineNum_**, **_StartingNum_**, **_FromText_**, **_CountBy_**, **_NumMode_**, **_TwoOnOne_**, **_GutterPosition_**, **_LayoutMode_**, **_CharsLine_**, **_LinesPage_**, **_CharPitch_**, **_LinePitch_**, **_DocFontName_**, **_DocFontSize_**, **_PageColumns_**, **_TextFlow_**, **_FirstPageOnLeft_**, **_SectionType_**, **_RTLAlignment_**|
| **wdDialogFileFind**| **_SearchName_** , **_SearchPath_**, **_Name_**, **_SubDir_**, **_Title_**, **_Author_**, **_Keywords_**, **_Subject_**, **_Options_**, **_MatchCase_**, **_Text_**, **_PatternMatch_**, **_DateSavedFrom_**, **_DateSavedTo_**, **_SavedBy_**, **_DateCreatedFrom_**, **_DateCreatedTo_**, **_View_**, **_SortBy_**, **_ListBy_**, **_SelectedFile_**, **_Add_**, **_Delete_**, **_ShowFolders_**, **_MatchByte_**|
| **wdDialogFileMacPageSetup**|(For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.)|
| **wdDialogFileNew**| **_Template_** , **_NewTemplate_**, **_DocumentType_**, **_Visible_**|
| **wdDialogFileOpen**| **_Name_** , **_ConfirmConversions_**, **_ReadOnly_**, **_LinkToSource_**, **_AddToMru_**, **_PasswordDoc_**, **_PasswordDot_**, **_Revert_**, **_WritePasswordDoc_**, **_WritePasswordDot_**, **_Connection_**, **_SQLStatement_**, **_SQLStatement1_**, **_Format_**, **_Encoding_**, **_Visible_**, **_OpenExclusive_**, **_OpenAndRepair_**, **_SubType_**, **_DocumentDirection_**, **_NoEncodingDialog_**, **_XMLTransform_**|
| **wdDialogFilePageSetup**| **_Tab_** , **_PaperSize_**, **_TopMargin_**, **_BottomMargin_**, **_LeftMargin_**, **_RightMargin_**, **_Gutter_**, **_PageWidth_**, **_PageHeight_**, **_Orientation_**, **_FirstPage_**, **_OtherPages_**, **_VertAlign_**, **_ApplyPropsTo_**, **_Default_**, **_FacingPages_**, **_HeaderDistance_**, **_FooterDistance_**, **_SectionStart_**, **_OddAndEvenPages_**, **_DifferentFirstPage_**, **_Endnotes_**, **_LineNum_**, **_StartingNum_**, **_FromText_**, **_CountBy_**, **_NumMode_**, **_TwoOnOne_**, **_GutterPosition_**, **_LayoutMode_**, **_CharsLine_**, **_LinesPage_**, **_CharPitch_**, **_LinePitch_**, **_DocFontName_**, **_DocFontSize_**, **_PageColumns_**, **_TextFlow_**, **_FirstPageOnLeft_**, **_SectionType_**, **_RTLAlignment_**, **_FolioPrint_**, **_ReverseFolio_**, **_FolioPages_**|
| **wdDialogFilePrint**| **_Background_** , **_AppendPrFile_**, **_Range_**, **_PrToFileName_**, **_From_**, **_To_**, **_Type_**, **_NumCopies_**, **_Pages_**, **_Order_**, **_PrintToFile_**, **_Collate_**, **_FileName_**, **_Printer_**, **_OutputPrinter_**, **_DuplexPrint_**, **_PrintZoomColumn_**, **_PrintZoomRow_**, **_PrintZoomPaperWidth_**, **_PrintZoomPaperHeight_**, **_ZoomPaper_**|
| **wdDialogFilePrintOneCopy**|(For information about this constant, consult the language reference Help included with Microsoft Office Macintosh Edition.)|
| **wdDialogFilePrintSetup**| **_Printer_** , **_Options_**, **_Network_**, **_DoNotSetAsSysDefault_**|
| **wdDialogFileRoutingSlip**| **_Subject_** , **_Message_**, **_AllAtOnce_**, **_ReturnWhenDone_**, **_TrackStatus_**, **_Protect_**, **_AddSlip_**, **_RouteDocument_**, **_AddRecipient_**, **_OldRecipient_**, **_ResetSlip_**, **_ClearSlip_**, **_ClearRecipients_**, **_Address_**|
| **wdDialogFileSaveAs**| **_Name_** , **_Format_**, **_LockAnnot_**, **_Password_**, **_AddToMru_**, **_WritePassword_**, **_RecommendReadOnly_**, **_EmbedFonts_**, **_NativePictureFormat_**, **_FormsData_**, **_SaveAsAOCELetter_**, **_WriteVersion_**, **_VersionDesc_**, **_InsertLineBreaks_**, **_AllowSubstitutions_**, **_LineEnding_**, **_AddBiDiMarks_**|
| **wdDialogFileSaveVersion**|(none)|
| **wdDialogFileSummaryInfo**| **_Title_** , **_Subject_**, **_Author_**, **_Keywords_**, **_Comments_**, **_FileName_**, **_Directory_**, **_Template_**, **_CreateDate_**, **_LastSavedDate_**, **_LastSavedBy_**, **_RevisionNumber_**, **_EditTime_**, **_LastPrintedDate_**, **_NumPages_**, **_NumWords_**, **_NumChars_**, **_NumParas_**, **_NumLines_**, **_Update_**, **_FileSize_**|
| **wdDialogFileVersions**| **_AutoVersion_** , **_VersionDesc_**|
| **wdDialogFitText**| **_FitTextWidth_**|
| **wdDialogFontSubstitution**| **_UnavailableFont_** , **_SubstituteFont_**|
| **wdDialogFormatAddrFonts**| **_Points_** , **_Underline_**, **_Color_**, **_StrikeThrough_**, **_Superscript_**, **_Subscript_**, **_Hidden_**, **_SmallCaps_**, **_AllCaps_**, **_Spacing_**, **_Position_**, **_Kerning_**, **_KerningMin_**, **_Default_**, **_Tab_**, **_Font_**, **_Bold_**, **_Italic_**, **_DoubleStrikeThrough_**, **_Shadow_**, **_Outline_**, **_Emboss_**, **_Engrave_**, **_Scale_**, **_Animations_**, **_CharAccent_**, **_FontMajor_**, **_FontLowAnsi_**, **_FontHighAnsi_**, **_CharacterWidthGrid_**, **_ColorRGB_**, **_UnderlineColor_**, **_PointsBi_**, **_ColorBi_**, **_FontNameBi_**, **_BoldBi_**, **_ItalicBi_**, **_DiacColor_**|
| **wdDialogFormatBordersAndShading**| **_ApplyTo_** , **_Shadow_**, **_TopBorder_**, **_LeftBorder_**, **_BottomBorder_**, **_RightBorder_**, **_HorizBorder_**, **_VertBorder_**, **_TopColor_**, **_LeftColor_**, **_BottomColor_**, **_RightColor_**, **_HorizColor_**, **_VertColor_**, **_FromText_**, **_Shading_**, **_Foreground_**, **_Background_**, **_Tab_**, **_FineShading_**, **_TopStyle_**, **_LeftStyle_**, **_BottomStyle_**, **_RightStyle_**, **_HorizStyle_**, **_VertStyle_**, **_TopWeight_**, **_LeftWeight_**, **_BottomWeight_**, **_RightWeight_**, **_HorizWeight_**, **_VertWeight_**, **_BorderObjectType_**, **_BorderArtWeight_**, **_BorderArt_**, **_FromTextTop_**, **_FromTextBottom_**, **_FromTextLeft_**, **_FromTextRight_**, **_OffsetFrom_**, **_InFront_**, **_SurroundHeader_**, **_SurroundFooter_**, **_JoinBorder_**, **_LineColor_**, **_WhichPages_**, **_TL2BRBorder_**, **_TR2BLBorder_**, **_TL2BRColor_**, **_TR2BLColor_**, **_TL2BRStyle_**, **_TR2BLStyle_**, **_TL2BRWeight_**, **_TR2BLWeight_**, **_ForegroundRGB_**, **_BackgroundRGB_**, **_TopColorRGB_**, **_LeftColorRGB_**, **_BottomColorRGB_**, **_RightColorRGB_**, **_HorizColorRGB_**, **_VertColorRGB_**, **_TL2BRColorRGB_**, **_TR2BLColorRGB_**, **_LineColorRGB_**|
| **wdDialogFormatBulletsAndNumbering**|(none)|
| **wdDialogFormatCallout**| **_Type_** , **_Gap_**, **_Angle_**, **_Drop_**, **_Length_**, **_Border_**, **_AutoAttach_**, **_Accent_**|
| **wdDialogFormatChangeCase**| **_Type_**|
| **wdDialogFormatColumns**| **_Columns_** , **_ColumnNo_**, **_ColumnWidth_**, **_ColumnSpacing_**, **_EvenlySpaced_**, **_ApplyColsTo_**, **_ColLine_**, **_StartNewCol_**, **_FlowColumnsRtl_**|
| **wdDialogFormatDefineStyleBorders**| **_ApplyTo_** , **_Shadow_**, **_TopBorder_**, **_LeftBorder_**, **_BottomBorder_**, **_RightBorder_**, **_HorizBorder_**, **_VertBorder_**, **_TopColor_**, **_LeftColor_**, **_BottomColor_**, **_RightColor_**, **_HorizColor_**, **_VertColor_**, **_FromText_**, **_Shading_**, **_Foreground_**, **_Background_**, **_Tab_**, **_FineShading_**, **_TopStyle_**, **_LeftStyle_**, **_BottomStyle_**, **_RightStyle_**, **_HorizStyle_**, **_VertStyle_**, **_TopWeight_**, **_LeftWeight_**, **_BottomWeight_**, **_RightWeight_**, **_HorizWeight_**, **_VertWeight_**, **_BorderObjectType_**, **_BorderArtWeight_**, **_BorderArt_**, **_FromTextTop_**, **_FromTextBottom_**, **_FromTextLeft_**, **_FromTextRight_**, **_OffsetFrom_**, **_InFront_**, **_SurroundHeader_**, **_SurroundFooter_**, **_JoinBorder_**, **_LineColor_**, **_WhichPages_**, **_TL2BRBorder_**, **_TR2BLBorder_**, **_TL2BRColor_**, **_TR2BLColor_**, **_TL2BRStyle_**, **_TR2BLStyle_**, **_TL2BRWeight_**, **_TR2BLWeight_**, **_ForegroundRGB_**, **_BackgroundRGB_**, **_TopColorRGB_**, **_LeftColorRGB_**, **_BottomColorRGB_**, **_RightColorRGB_**, **_HorizColorRGB_**, **_VertColorRGB_**, **_TL2BRColorRGB_**, **_TR2BLColorRGB_**, **_LineColorRGB_**|
| **wdDialogFormatDefineStyleFont**| **_Points_** , **_Underline_**, **_Color_**, **_StrikeThrough_**, **_Superscript_**, **_Subscript_**, **_Hidden_**, **_SmallCaps_**, **_AllCaps_**, **_Spacing_**, **_Position_**, **_Kerning_**, **_KerningMin_**, **_Default_**, **_Tab_**, **_Font_**, **_Bold_**, **_Italic_**, **_DoubleStrikeThrough_**, **_Shadow_**, **_Outline_**, **_Emboss_**, **_Engrave_**, **_Scale_**, **_Animations_**, **_CharAccent_**, **_FontMajor_**, **_FontLowAnsi_**, **_FontHighAnsi_**, **_CharacterWidthGrid_**, **_ColorRGB_**, **_UnderlineColor_**, **_PointsBi_**, **_ColorBi_**, **_FontNameBi_**, **_BoldBi_**, **_ItalicBi_**, **_DiacColor_**|
| **wdDialogFormatDefineStyleFrame**| **_Wrap_** , **_WidthRule_**, **_FixedWidth_**, **_HeightRule_**, **_FixedHeight_**, **_PositionHorz_**, **_PositionHorzRel_**, **_DistFromText_**, **_PositionVert_**, **_PositionVertRel_**, **_DistVertFromText_**, **_MoveWithText_**, **_LockAnchor_**, **_RemoveFrame_**|
| **wdDialogFormatDefineStyleLang**| **_Language_** , **_CheckLanguage_**, **_Default_**, **_NoProof_**|
| **wdDialogFormatDefineStylePara**| **_LeftIndent_** , **_RightIndent_**, **_Before_**, **_After_**, **_LineSpacingRule_**, **_LineSpacing_**, **_Alignment_**, **_WidowControl_**, **_KeepWithNext_**, **_KeepTogether_**, **_PageBreak_**, **_NoLineNum_**, **_DontHyphen_**, **_Tab_**, **_FirstIndent_**, **_OutlineLevel_**, **_Kinsoku_**, **_WordWrap_**, **_OverflowPunct_**, **_TopLinePunct_**, **_AutoSpaceDE_**, **_LineHeightGrid_**, **_AutoSpaceDN_**, **_CharAlign_**, **_CharacterUnitLeftIndent_**, **_AdjustRight_**, **_CharacterUnitFirstIndent_**, **_CharacterUnitRightIndent_**, **_LineUnitBefore_**, **_LineUnitAfter_**, **_NoSpaceBetweenParagraphsOfSameStyle_**, **_OrientationBi_**|
| **wdDialogFormatDefineStyleTabs**| **_Position_** , **_DefTabs_**, **_Align_**, **_Leader_**, **_Set_**, **_Clear_**, **_ClearAll_**|
| **wdDialogFormatDrawingObject**| **_Left_** , **_PositionHorzRel_**, **_Top_**, **_PositionVertRel_**, **_LockAnchor_**, **_FloatOverText_**, **_LayoutInCell_**, **_WrapSide_**, **_TopDistanceFromText_**, **_BottomDistanceFromText_**, **_LeftDistanceFromText_**, **_RightDistanceFromText_**, **_Wrap_**, **_WordWrap_**, **_AutoSize_**, **_HRWidthType_**, **_HRHeight_**, **_HRNoshade_**, **_HRAlign_**, **_Text_**, **_AllowOverlap_**, **_HorizRule_**|
| **wdDialogFormatDropCap**| **_Position_** , **_Font_**, **_DropHeight_**, **_DistFromText_**|
| **wdDialogFormatEncloseCharacters**| **_Style_** , **_Text_**, **_Enclosure_**|
| **wdDialogFormatFont**| **_Points_** , **_Underline_**, **_Color_**, **_StrikeThrough_**, **_Superscript_**, **_Subscript_**, **_Hidden_**, **_SmallCaps_**, **_AllCaps_**, **_Spacing_**, **_Position_**, **_Kerning_**, **_KerningMin_**, **_Default_**, **_Tab_**, **_Font_**, **_Bold_**, **_Italic_**, **_DoubleStrikeThrough_**, **_Shadow_**, **_Outline_**, **_Emboss_**, **_Engrave_**, **_Scale_**, **_Animations_**, **_CharAccent_**, **_FontMajor_**, **_FontLowAnsi_**, **_FontHighAnsi_**, **_CharacterWidthGrid_**, **_ColorRGB_**, **_UnderlineColor_**, **_PointsBi_**, **_ColorBi_**, **_FontNameBi_**, **_BoldBi_**, **_ItalicBi_**, **_DiacColor_**|
| **wdDialogFormatFrame**| **_Wrap_** , **_WidthRule_**, **_FixedWidth_**, **_HeightRule_**, **_FixedHeight_**, **_PositionHorz_**, **_PositionHorzRel_**, **_DistFromText_**, **_PositionVert_**, **_PositionVertRel_**, **_DistVertFromText_**, **_MoveWithText_**, **_LockAnchor_**, **_RemoveFrame_**|
| **wdDialogFormatPageNumber**| **_ChapterNumber_** , **_NumRestart_**, **_NumFormat_**, **_StartingNum_**, **_Level_**, **_Separator_**, **_DoubleQuote_**, **_PgNumberingStyle_**|
| **wdDialogFormatParagraph**| **_LeftIndent_** , **_RightIndent_**, **_Before_**, **_After_**, **_LineSpacingRule_**, **_LineSpacing_**, **_Alignment_**, **_WidowControl_**, **_KeepWithNext_**, **_KeepTogether_**, **_PageBreak_**, **_NoLineNum_**, **_DontHyphen_**, **_Tab_**, **_FirstIndent_**, **_OutlineLevel_**, **_Kinsoku_**, **_WordWrap_**, **_OverflowPunct_**, **_TopLinePunct_**, **_AutoSpaceDE_**, **_LineHeightGrid_**, **_AutoSpaceDN_**, **_CharAlign_**, **_CharacterUnitLeftIndent_**, **_AdjustRight_**, **_CharacterUnitFirstIndent_**, **_CharacterUnitRightIndent_**, **_LineUnitBefore_**, **_LineUnitAfter_**, **_NoSpaceBetweenParagraphsOfSameStyle_**, **_OrientationBi_**|
| **wdDialogFormatPicture**| **_SetSize_** , **_CropLeft_**, **_CropRight_**, **_CropTop_**, **_CropBottom_**, **_ScaleX_**, **_ScaleY_**, **_SizeX_**, **_SizeY_**|
| **wdDialogFormatRetAddrFonts**| **_Points_** , **_Underline_**, **_Color_**, **_StrikeThrough_**, **_Superscript_**, **_Subscript_**, **_Hidden_**, **_SmallCaps_**, **_AllCaps_**, **_Spacing_**, **_Position_**, **_Kerning_**, **_KerningMin_**, **_Default_**, **_Tab_**, **_Font_**, **_Bold_**, **_Italic_**, **_DoubleStrikeThrough_**, **_Shadow_**, **_Outline_**, **_Emboss_**, **_Engrave_**, **_Scale_**, **_Animations_**, **_CharAccent_**, **_FontMajor_**, **_FontLowAnsi_**, **_FontHighAnsi_**, **_CharacterWidthGrid_**, **_ColorRGB_**, **_UnderlineColor_**, **_PointsBi_**, **_ColorBi_**, **_FontNameBi_**, **_BoldBi_**, **_ItalicBi_**, **_DiacColor_**|
| **wdDialogFormatSectionLayout**| **_SectionStart_** , **_VertAlign_**, **_Endnotes_**, **_LineNum_**, **_StartingNum_**, **_FromText_**, **_CountBy_**, **_NumMode_**, **_SectionType_**|
| **wdDialogFormatStyle**| **_Name_** , **_Delete_**, **_Merge_**, **_NewName_**, **_BasedOn_**, **_NextStyle_**, **_Type_**, **_FileName_**, **_Source_**, **_AddToTemplate_**, **_Define_**, **_Rename_**, **_Apply_**, **_New_**, **_Link_**|
| **wdDialogFormatStyleGallery**| **_Template_** , **_Preview_**|
| **wdDialogFormatStylesCustom**|(none)|
| **wdDialogFormatTabs**| **_Position_** , **_DefTabs_**, **_Align_**, **_Leader_**, **_Set_**, **_Clear_**, **_ClearAll_**|
| **wdDialogFormatTheme**|(none)|
| **wdDialogFormFieldHelp**|(none)|
| **wdDialogFormFieldOptions**| **_Entry_** , **_Exit_**, **_Name_**, **_Enable_**, **_TextType_**, **_TextWidth_**, **_TextDefault_**, **_TextFormat_**, **_CheckSize_**, **_CheckWidth_**, **_CheckDefault_**, **_Type_**, **_OwnHelp_**, **_HelpText_**, **_OwnStat_**, **_StatText_**, **_Calculate_**|
| **wdDialogFrameSetProperties**|(none)|
| **wdDialogHelpAbout**| **_APPNAME_** , **_APPCOPYRIGHT_**, **_APPUSERNAME_**, **_APPORGANIZATION_**, **_APPSERIALNUMBER_**|
| **wdDialogHelpWordPerfectHelp**| **_WPCommand_** , **_HelpText_**, **_DemoGuidance_**|
| **wdDialogHelpWordPerfectHelpOptions**| **_CommandKeyHelp_** , **_DocNavKeys_**, **_MouseSimulation_**, **_DemoGuidance_**, **_DemoSpeed_**, **_HelpType_**|
| **wdDialogHorizontalInVertical**|(none)|
| **wdDialogIMESetDefault**|(none)|
| **wdDialogInsertAddCaption**| **_Name_**|
| **wdDialogInsertAutoCaption**| **_Clear_** , **_ClearAll_**, **_Object_**, **_Label_**, **_Position_**|
| **wdDialogInsertBookmark**| **_Name_** , **_SortBy_**, **_Add_**, **_Delete_**, **_Goto_**, **_Hidden_**|
| **wdDialogInsertBreak**| **_Type_**|
| **wdDialogInsertCaption**| **_Label_** , **_TitleAutoText_**, **_Title_**, **_Delete_**, **_Position_**, **_AutoCaption_**, **_ExcludeLabel_**|
| **wdDialogInsertCaptionNumbering**| **_Label_** , **_FormatNumber_**, **_ChapterNumber_**, **_Level_**, **_Separator_**, **_CapNumberingStyle_**|
| **wdDialogInsertCrossReference**| **_ReferenceType_** , **_ReferenceKind_**, **_ReferenceItem_**, **_InsertAsHyperLink_**, **_InsertPosition_**, **_SeparateNumbers_**, **_SeparatorCharacters_**|
| **wdDialogInsertDatabase**| **_Format_** , **_Style_**, **_LinkToSource_**, **_Connection_**, **_SQLStatement_**, **_SQLStatement1_**, **_PasswordDoc_**, **_PasswordDot_**, **_DataSource_**, **_From_**, **_To_**, **_IncludeFields_**, **_WritePasswordDoc_**, **_WritePasswordDot_**|
| **wdDialogInsertDateTime**| **_DateTimePic_** , **_InsertAsField_**, **_DbCharField_**, **_DateLanguage_**, **_CalendarType_**|
| **wdDialogInsertField**| **_Field_**|
| **wdDialogInsertFile**| **_Name_** , **_Range_**, **_ConfirmConversions_**, **_Link_**, **_Attachment_**|
| **wdDialogInsertFootnote**| **_Reference_** , **_NoteType_**, **_Symbol_**, **_FootNumberAs_**, **_EndNumberAs_**, **_FootnotesAt_**, **_EndnotesAt_**, **_FootNumberingStyle_**, **_EndNumberingStyle_**, **_FootStartingNum_**, **_FootRestartNum_**, **_EndStartingNum_**, **_EndRestartNum_**, **_ApplyPropsTo_**|
| **wdDialogInsertFormField**| **_Entry_** , **_Exit_**, **_Name_**, **_Enable_**, **_TextType_**, **_TextWidth_**, **_TextDefault_**, **_TextFormat_**, **_CheckSize_**, **_CheckWidth_**, **_CheckDefault_**, **_Type_**, **_OwnHelp_**, **_HelpText_**, **_OwnStat_**, **_StatText_**, **_Calculate_**|
| **wdDialogInsertHyperlink**|(none)|
| **wdDialogInsertIndex**| **_Outline_** , **_Fields_**, **_From_**, **_To_**, **_TableId_**, **_AddedStyles_**, **_Caption_**, **_HeadingSeparator_**, **_Replace_**, **_MarkEntry_**, **_AutoMark_**, **_MarkCitation_**, **_Type_**, **_RightAlignPageNumbers_**, **_Passim_**, **_KeepFormatting_**, **_Columns_**, **_Category_**, **_Label_**, **_ShowPageNumbers_**, **_AccentedLetters_**, **_Filter_**, **_SortBy_**, **_Leader_**, **_TOCUseHyperlinks_**, **_TOCHidePageNumInWeb_**, **_IndexLanguage_**, **_UseOutlineLevel_**|
| **wdDialogInsertIndexAndTables**| **_Outline_** , **_Fields_**, **_From_**, **_To_**, **_TableId_**, **_AddedStyles_**, **_Caption_**, **_HeadingSeparator_**, **_Replace_**, **_MarkEntry_**, **_AutoMark_**, **_MarkCitation_**, **_Type_**, **_RightAlignPageNumbers_**, **_Passim_**, **_KeepFormatting_**, **_Columns_**, **_Category_**, **_Label_**, **_ShowPageNumbers_**, **_AccentedLetters_**, **_Filter_**, **_SortBy_**, **_Leader_**, **_TOCUseHyperlinks_**, **_TOCHidePageNumInWeb_**, **_IndexLanguage_**, **_UseOutlineLevel_**|
| **wdDialogInsertMergeField**| **_MergeField_** , **_WordField_**|
| **wdDialogInsertNumber**| **_NumPic_**|
| **wdDialogInsertObject**| **_IconNumber_** , **_FileName_**, **_Link_**, **_DisplayIcon_**, **_Tab_**, **_Class_**, **_IconFileName_**, **_Caption_**, **_Floating_**|
| **wdDialogInsertPageNumbers**| **_Type_** , **_Position_**, **_FirstPage_**|
| **wdDialogInsertPicture**| **_Name_** , **_LinkToFile_**, **_New_**, **_FloatOverText_**|
| **wdDialogInsertSubdocument**| **_Name_** , **_ConfirmConversions_**, **_ReadOnly_**, **_LinkToSource_**, **_AddToMru_**, **_PasswordDoc_**, **_PasswordDot_**, **_Revert_**, **_WritePasswordDoc_**, **_WritePasswordDot_**, **_Connection_**, **_SQLStatement_**, **_SQLStatement1_**, **_Format_**, **_Encoding_**, **_Visible_**, **_OpenExclusive_**, **_OpenAndRepair_**, **_SubType_**, **_DocumentDirection_**, **_NoEncodingDialog_**, **_XMLTransform_**|
| **wdDialogInsertSymbol**| **_Font_** , **_Tab_**, **_CharNum_**, **_CharNumLow_**, **_Unicode_**, **_Hint_**|
| **wdDialogInsertTableOfAuthorities**| **_Outline_** , **_Fields_**, **_From_**, **_To_**, **_TableId_**, **_AddedStyles_**, **_Caption_**, **_HeadingSeparator_**, **_Replace_**, **_MarkEntry_**, **_AutoMark_**, **_MarkCitation_**, **_Type_**, **_RightAlignPageNumbers_**, **_Passim_**, **_KeepFormatting_**, **_Columns_**, **_Category_**, **_Label_**, **_ShowPageNumbers_**, **_AccentedLetters_**, **_Filter_**, **_SortBy_**, **_Leader_**, **_TOCUseHyperlinks_**, **_TOCHidePageNumInWeb_**, **_IndexLanguage_**, **_UseOutlineLevel_**|
| **wdDialogInsertTableOfContents**| **_Outline_** , **_Fields_**, **_From_**, **_To_**, **_TableId_**, **_AddedStyles_**, **_Caption_**, **_HeadingSeparator_**, **_Replace_**, **_MarkEntry_**, **_AutoMark_**, **_MarkCitation_**, **_Type_**, **_RightAlignPageNumbers_**, **_Passim_**, **_KeepFormatting_**, **_Columns_**, **_Category_**, **_Label_**, **_ShowPageNumbers_**, **_AccentedLetters_**, **_Filter_**, **_SortBy_**, **_Leader_**, **_TOCUseHyperlinks_**, **_TOCHidePageNumInWeb_**, **_IndexLanguage_**, **_UseOutlineLevel_**|
| **wdDialogInsertTableOfFigures**| **_Outline_** , **_Fields_**, **_From_**, **_To_**, **_TableId_**, **_AddedStyles_**, **_Caption_**, **_HeadingSeparator_**, **_Replace_**, **_MarkEntry_**, **_AutoMark_**, **_MarkCitation_**, **_Type_**, **_RightAlignPageNumbers_**, **_Passim_**, **_KeepFormatting_**, **_Columns_**, **_Category_**, **_Label_**, **_ShowPageNumbers_**, **_AccentedLetters_**, **_Filter_**, **_SortBy_**, **_Leader_**, **_TOCUseHyperlinks_**, **_TOCHidePageNumInWeb_**, **_IndexLanguage_**, **_UseOutlineLevel_**|
| **wdDialogInsertWebComponent**| **_IconNumber_** , **_FileName_**, **_Link_**, **_DisplayIcon_**, **_Tab_**, **_Class_**, **_IconFileName_**, **_Caption_**, **_Floating_**|
| **wdDialogLetterWizard**| **_SenderCity_** , **_DateFormat_**, **_IncludeHeaderFooter_**, **_LetterStyle_**, **_Letterhead_**, **_LetterheadLocation_**, **_LetterheadSize_**, **_RecipientName_**, **_RecipientAddress_**, **_Salutation_**, **_SalutationType_**, **_RecipientGender_**, **_RecipientReference_**, **_MailingInstructions_**, **_AttentionLine_**, **_LetterSubject_**, **_CCList_**, **_SenderName_**, **_ReturnAddress_**, **_Closing_**, **_SenderJobTitle_**, **_SenderCompany_**, **_SenderInitials_**, **_EnclosureNumber_**, **_PageDesign_**, **_InfoBlock_**, **_SenderGender_**, **_ReturnAddressSF_**, **_RecipientCode_**, **_SenderCode_**, **_SenderReference_**|
| **wdDialogListCommands**| **_ListType_**|
| **wdDialogMailMerge**| **_CheckErrors_** , **_Destination_**, **_MergeRecords_**, **_From_**, **_To_**, **_Suppression_**, **_MailMerge_**, **_QueryOptions_**, **_MailSubject_**, **_MailAsAttachment_**, **_MailAddress_**|
| **wdDialogMailMergeCheck**| **_CheckErrors_**|
| **wdDialogMailMergeCreateDataSource**| **_FileName_** , **_PasswordDoc_**, **_PasswordDot_**, **_HeaderRecord_**, **_MSQuery_**, **_SQLStatement_**, **_SQLStatement1_**, **_Connection_**, **_LinkToSource_**, **_WritePasswordDoc_**|
| **wdDialogMailMergeCreateHeaderSource**| **_FileName_** , **_PasswordDoc_**, **_PasswordDot_**, **_HeaderRecord_**, **_MSQuery_**, **_SQLStatement_**, **_SQLStatement1_**, **_Connection_**, **_LinkToSource_**, **_WritePasswordDoc_**|
| **wdDialogMailMergeFieldMapping**|(none)|
| **wdDialogMailMergeFindRecipient**|(none)|
| **wdDialogMailMergeFindRecord**| **_Find_** , **_Field_**|
| **wdDialogMailMergeHelper**| **_Merge_** , **_Options_**|
| **wdDialogMailMergeInsertAddressBlock**|(none)|
| **wdDialogMailMergeInsertAsk**| **_Name_** , **_Prompt_**, **_DefaultBookmarkText_**, **_AskOnce_**|
| **wdDialogMailMergeInsertFields**|(none)|
| **wdDialogMailMergeInsertFillIn**| **_Prompt_** , **_DefaultFillInText_**, **_AskOnce_**|
| **wdDialogMailMergeInsertGreetingLine**|(none)|
| **wdDialogMailMergeInsertIf**| **_MergeField_** , **_Comparison_**, **_CompareTo_**, **_TrueAutoText_**, **_TrueText_**, **_FalseAutoText_**, **_FalseText_**|
| **wdDialogMailMergeInsertNextIf**| **_MergeField_** , **_Comparison_**, **_CompareTo_**|
| **wdDialogMailMergeInsertSet**| **_Name_** , **_ValueText_**, **_ValueAutoText_**|
| **wdDialogMailMergeInsertSkipIf**| **_MergeField_** , **_Comparison_**, **_CompareTo_**|
| **wdDialogMailMergeOpenDataSource**| **_Name_** , **_ConfirmConversions_**, **_ReadOnly_**, **_LinkToSource_**, **_AddToMru_**, **_PasswordDoc_**, **_PasswordDot_**, **_Revert_**, **_WritePasswordDoc_**, **_WritePasswordDot_**, **_Connection_**, **_SQLStatement_**, **_SQLStatement1_**, **_Format_**, **_Encoding_**, **_Visible_**, **_OpenExclusive_**, **_OpenAndRepair_**, **_SubType_**, **_DocumentDirection_**, **_NoEncodingDialog_**, **_XMLTransform_**|
| **wdDialogMailMergeOpenHeaderSource**| **_Name_** , **_ConfirmConversions_**, **_ReadOnly_**, **_LinkToSource_**, **_AddToMru_**, **_PasswordDoc_**, **_PasswordDot_**, **_Revert_**, **_WritePasswordDoc_**, **_WritePasswordDot_**, **_Connection_**, **_SQLStatement_**, **_SQLStatement1_**, **_Format_**, **_Encoding_**, **_Visible_**, **_OpenExclusive_**, **_OpenAndRepair_**, **_SubType_**, **_DocumentDirection_**, **_NoEncodingDialog_**, **_XMLTransform_**|
| **wdDialogMailMergeQueryOptions**| **_SQLStatement_** , **_SQLStatement1_**|
| **wdDialogMailMergeRecipients**|(none)|
| **wdDialogMailMergeSetDocumentType**| **_Type_**|
| **wdDialogMailMergeUseAddressBook**| **_AddressBookType_**|
| **wdDialogMarkCitation**| **_LongCitation_** , **_LongCitationAutoText_**, **_Category_**, **_ShortCitation_**, **_NextCitation_**, **_Mark_**, **_MarkAll_**|
| **wdDialogMarkIndexEntry**| **_MarkAll_** , **_Entry_**, **_Range_**, **_Bold_**, **_Italic_**, **_CrossReference_**, **_EntryAutoText_**, **_CrossReferenceAutoText_**, **_Yomi_**|
| **wdDialogMarkTableOfContentsEntry**| **_Entry_** , **_EntryAutoText_**, **_TableId_**, **_Level_**|
| **wdDialogNewToolbar**| **_Name_** , **_Context_**|
| **wdDialogNoteOptions**| **_FootnotesAt_** , **_FootNumberAs_**, **_FootStartingNum_**, **_FootRestartNum_**, **_EndnotesAt_**, **_EndNumberAs_**, **_EndStartingNum_**, **_EndRestartNum_**, **_FootNumberingStyle_**, **_EndNumberingStyle_**|
| **wdDialogOrganizer**| **_Copy_** , **_Delete_**, **_Rename_**, **_Source_**, **_Destination_**, **_Name_**, **_NewName_**, **_Tab_**|
| **wdDialogPhoneticGuide**|(none)|
| **wdDialogReviewAfmtRevisions**|(none)|
| **wdDialogSearch**|(none)|
| **wdDialogShowRepairs**| **_Name_** , **_SortBy_**, **_Add_**, **_Delete_**, **_GoTo_**, **_Hidden_**|
| **wdDialogTableAutoFormat**| **_HideAutoFit_** , **_Preview_**, **_Format_**, **_Borders_**, **_Shading_**, **_Font_**, **_Color_**, **_AutoFit_**, **_HeadingRows_**, **_FirstColumn_**, **_LastRow_**, **_LastColumn_**|
| **wdDialogTableCellOptions**|(none)|
| **wdDialogTableColumnWidth**| **_ColumnWidth_** , **_SpaceBetweenCols_**, **_PrevColumn_**, **_NextColumn_**, **_AutoFit_**, **_RulerStyle_**|
| **wdDialogTableDeleteCells**| **_ShiftCells_**|
| **wdDialogTableFormatCell**| **_Category_**|
| **wdDialogTableFormula**| **_Formula_** , **_NumFormat_**|
| **wdDialogTableInsertCells**| **_ShiftCells_**|
| **wdDialogTableInsertRow**| **_NumRows_**|
| **wdDialogTableInsertTable**| **_ConvertFrom_** , **_NumColumns_**, **_NumRows_**, **_InitialColWidth_**, **_Wizard_**, **_Format_**, **_Apply_**, **_AutoFit_**, **_SetDefault_**, **_Word8_**, **_Style_**|
| **wdDialogTableOfCaptionsOptions**|(none)|
| **wdDialogTableOfContentsOptions**|(none)|
| **wdDialogTableProperties**| **_TableDirection_**|
| **wdDialogTableRowHeight**| **_RulerStyle_** , **_LineSpacingRule_**, **_LineSpacing_**, **_LeftIndent_**, **_AllowRowSplit_**, **_Alignment_**, **_PrevRow_**, **_NextRow_**, **_TableDir_**|
| **wdDialogTableSort**| **_DontSortHdr_** , **_FieldNum_**, **_Type_**, **_Order_**, **_FieldNum2_**, **_Type2_**, **_Order2_**, **_FieldNum3_**, **_Type3_**, **_Order3_**, **_Separator_**, **_SortColumn_**, **_CaseSensitive_**, **_SortBiDi_**, **_IgnoreHe_**, **_Diacritics_**, **_IgnoreThe_**, **_Kashida_**, **_Language_**, **_UsingNum_**, **_UsingNum2_**, **_UsingNum3_**|
| **wdDialogTableSplitCells**| **_NumColumns_** , **_NumRows_**, **_MergeBeforeSplit_**|
| **wdDialogTableTableOptions**|(none)|
| **wdDialogTableToText**| **_ConvertTo_** , **_NestedTables_**|
| **wdDialogTableWrapping**| **_PositionHorz_** , **_PositionHorzRel_**, **_PositionVert_**, **_PositionVertRel_**, **_TopDistanceFromText_**, **_BottomDistanceFromText_**, **_LeftDistanceFromText_**, **_RightDistanceFromText_**, **_MoveWithText_**, **_AllowOverlap_**, |
| **wdDialogTCSCTranslator**| **_Direction_** , **_Varients_**, **_TranslateCommon_**|
| **wdDialogTextToTable**| **_ConvertFrom_** , **_NumColumns_**, **_NumRows_**, **_InitialColWidth_**, **_Wizard_**, **_Format_**, **_Apply_**, **_AutoFit_**, **_SetDefault_**, **_Word8_**, **_Style_**|
| **wdDialogToolsAcceptRejectChanges**| **_ShowMarks_** , **_HideMarks_**, **_Wrap_**, **_FindPrevious_**, **_FindNext_**, **_AcceptRevisions_**, **_RejectRevisions_**, **_AcceptAll_**, **_RejectAll_**|
| **wdDialogToolsAdvancedSettings**| **_Application_** , **_Option_**, **_Setting_**, **_Delete_**, **_Set_**|
| **wdDialogToolsAutoCorrect**| **_ShowFineTuner_** , **_CapTable_**, **_InWordMail_**, **_InitialCaps_**, **_SentenceCaps_**, **_Days_**, **_CapsLock_**, **_ReplaceText_**, **_Formatting_**, **_Replace_**, **_With_**, **_Add_**, **_Delete_**, **_SmartQuotes_**, **_CorrectHangulAndAlphabet_**, **_ConvBrackets_**, **_ConvQuotes_**, **_ConvPunct_**, **_ReplaceTextFromSpellingChecker_**|
| **wdDialogToolsAutoCorrectExceptions**| **_Tab_** , **_Name_**, **_AutoAdd_**, **_Add_**, **_Delete_**|
| **wdDialogToolsAutoManager**| **_Tab_**|
| **wdDialogToolsAutoSummarize**| **_TextSize_** , **_Show_**, **_Update_**|
| **wdDialogToolsBulletsNumbers**| **_Replace_** , **_Font_**, **_CharNum_**, **_Type_**, **_FormatOutline_**, **_AutoUpdate_**, **_FormatNumber_**, **_Punctuation_**, **_StartAt_**, **_Points_**, **_Hang_**, **_Indent_**, **_Remove_**, **_DoubleQuote_**|
| **wdDialogToolsCompareDocuments**| **_CompareDestination_** , **_DetectFormatting_**, **_IgnoreCompareWarn_**, **_UseFormatFrom_**, **_AddToMru_**, **_Merge_**, **_FilterPrivacy_**, **_FilterDateAndTime_**, **_Name_**, **_CompareAuthor_**|
| **wdDialogToolsCreateDirectory**| **_Directory_**|
| **wdDialogToolsCreateEnvelope**| **_ExtractAddress_** , **_LabelListIndex_**, **_LabelIndex_**, **_LabelDotMatrix_**, **_LabelTray_**, **_LabelAcross_**, **_LabelDown_**, **_EnvAddress_**, **_EnvOmitReturn_**, **_EnvReturn_**, **_PrintBarCode_**, **_SingleLabel_**, **_LabelRow_**, **_LabelColumn_**, **_PrintEnvLabel_**, **_AddToDocument_**, **_EnvWidth_**, **_EnvHeight_**, **_EnvPaperSize_**, **_PrintFIMA_**, **_UseEnvFeeder_**, **_Tab_**, **_AddrAutoText_**, **_AddrText_**, **_AddrFromLeft_**, **_AddrFromTop_**, **_RetAddrFromLeft_**, **_RetAddrFromTop_**, **_LabelTopMargin_**, **_LabelSideMargin_**, **_LabelVertPitch_**, **_LabelHorPitch_**, **_LabelHeight_**, **_LabelWidth_**, **_CustomName_**, **_RetAddrText_**, **_EnvPaperName_**, **_DefaultFaceUp_**, **_DefaultOrientation_**, **_RetAddrAutoText_**, **_VerticalEnvelope_**, **_VerticalLabel_**, **_RecipientNamefromLeft_**, **_RecipientNamefromTop_**, **_RecipientPostalfromLeft_**, **_RecipientPostalfromTop_**, **_SenderNamefromLeft_**, **_SenderNamefromTop_**, **_SenderPostalfromLeft_**, **_SenderPostalfromTop_**, **_PrintEPostage_**, **_PrintEPostageLabel_**|
| **wdDialogToolsCreateLabels**| **_ExtractAddress_** , **_LabelListIndex_**, **_LabelIndex_**, **_LabelDotMatrix_**, **_LabelTray_**, **_LabelAcross_**, **_LabelDown_**, **_EnvAddress_**, **_EnvOmitReturn_**, **_EnvReturn_**, **_PrintBarCode_**, **_SingleLabel_**, **_LabelRow_**, **_LabelColumn_**, **_PrintEnvLabel_**, **_AddToDocument_**, **_EnvWidth_**, **_EnvHeight_**, **_EnvPaperSize_**, **_PrintFIMA_**, **_UseEnvFeeder_**, **_Tab_**, **_AddrAutoText_**, **_AddrText_**, **_AddrFromLeft_**, **_AddrFromTop_**, **_RetAddrFromLeft_**, **_RetAddrFromTop_**, **_LabelTopMargin_**, **_LabelSideMargin_**, **_LabelVertPitch_**, **_LabelHorPitch_**, **_LabelHeight_**, **_LabelWidth_**, **_CustomName_**, **_RetAddrText_**, **_EnvPaperName_**, **_DefaultFaceUp_**, **_DefaultOrientation_**, **_RetAddrAutoText_**, **_VerticalEnvelope_**, **_VerticalLabel_**, **_RecipientNamefromLeft_**, **_RecipientNamefromTop_**, **_RecipientPostalfromLeft_**, **_RecipientPostalfromTop_**, **_SenderNamefromLeft_**, **_SenderNamefromTop_**, **_SenderPostalfromLeft_**, **_SenderPostalfromTop_**, **_PrintEPostage_**, **_PrintEPostageLabel_**|
| **wdDialogToolsCustomize**| **_KeyCode_** , **_KeyCode2_**, **_MenuType_**, **_Position_**, **_AddAll_**, **_Category_**, **_Name_**, **_Menu_**, **_AddBelow_**, **_MenuText_**, **_Rename_**, **_Add_**, **_Remove_**, **_ResetAll_**, **_CommandValue_**, **_Context_**, **_Tab_**|
| **wdDialogToolsCustomizeKeyboard**| **_KeyCode_** , **_KeyCode2_**, **_MenuType_**, **_Position_**, **_AddAll_**, **_Category_**, **_Name_**, **_Menu_**, **_AddBelow_**, **_MenuText_**, **_Rename_**, **_Add_**, **_Remove_**, **_ResetAll_**, **_CommandValue_**, **_Context_**, **_Tab_**|
| **wdDialogToolsCustomizeMenuBar**| **_Context_** , **_Position_**, **_MenuType_**, **_MenuText_**, **_Menu_**, **_Add_**, **_Remove_**, **_Rename_**|
| **wdDialogToolsCustomizeMenus**| **_KeyCode_** , **_KeyCode2_**, **_MenuType_**, **_Position_**, **_AddAll_**, **_Category_**, **_Name_**, **_Menu_**, **_AddBelow_**, **_MenuText_**, **_Rename_**, **_Add_**, **_Remove_**, **_ResetAll_**, **_CommandValue_**, **_Context_**, **_Tab_**|
| **wdDialogToolsEnvelopesAndLabels**| **_ExtractAddress_** , **_LabelListIndex_**, **_LabelIndex_**, **_LabelDotMatrix_**, **_LabelTray_**, **_LabelAcross_**, **_LabelDown_**, **_EnvAddress_**, **_EnvOmitReturn_**, **_EnvReturn_**, **_PrintBarCode_**, **_SingleLabel_**, **_LabelRow_**, **_LabelColumn_**, **_PrintEnvLabel_**, **_AddToDocument_**, **_EnvWidth_**, **_EnvHeight_**, **_EnvPaperSize_**, **_PrintFIMA_**, **_UseEnvFeeder_**, **_Tab_**, **_AddrAutoText_**, **_AddrText_**, **_AddrFromLeft_**, **_AddrFromTop_**, **_RetAddrFromLeft_**, **_RetAddrFromTop_**, **_LabelTopMargin_**, **_LabelSideMargin_**, **_LabelVertPitch_**, **_LabelHorPitch_**, **_LabelHeight_**, **_LabelWidth_**, **_CustomName_**, **_RetAddrText_**, **_EnvPaperName_**, **_DefaultFaceUp_**, **_DefaultOrientation_**, **_RetAddrAutoText_**, **_VerticalEnvelope_**, **_VerticalLabel_**, **_RecipientNamefromLeft_**, **_RecipientNamefromTop_**, **_RecipientPostalfromLeft_**, **_RecipientPostalfromTop_**, **_SenderNamefromLeft_**, **_SenderNamefromTop_**, **_SenderPostalfromLeft_**, **_SenderPostalfromTop_**, **_PrintEPostage_**, **_PrintEPostageLabel_**|
| **wdDialogToolsGrammarSettings**| **_Options_**|
| **wdDialogToolsHangulHanjaConversion**|(none)|
| **wdDialogToolsHighlightChanges**| **_MarkRevisions_** , **_ViewRevisions_**, **_PrintRevisions_**, **_AcceptAll_**, **_RejectAll_**|
| **wdDialogToolsHyphenation**| **_AutoHyphenation_** , **_HyphenateCaps_**, **_HyphenationZone_**, **_LimitConsecutiveHyphens_**|
| **wdDialogToolsLanguage**| **_Language_** , **_CheckLanguage_**, **_Default_**, **_NoProof_**|
| **wdDialogToolsMacro**| **_Name_** , **_Run_**, **_Edit_**, **_Show_**, **_Delete_**, **_Rename_**, **_Description_**, **_NewName_**, **_SetDesc_**|
| **wdDialogToolsMacroRecord**|(This dialog box cannot be called from a macro.)|
| **wdDialogToolsManageFields**| **_FieldName_** , **_Add_**, **_Remove_**, **_Rename_**, **_NewName_**|
| **wdDialogToolsMergeDocuments**| **_Name_**|
| **wdDialogToolsOptions**| **_Tab_**|
| **wdDialogToolsOptionsAutoFormat**| **_ApplyStylesLists_** , **_ApplyBulletedLists_**, **_ApplyStylesOtherParas_**, **_ReplaceQuotes_**, **_ReplaceOrdinals_**, **_ReplaceFractions_**, **_ReplaceSymbols_**, **_ReplacePlainTextEmphasis_**, **_ReplaceHyperlinks_**, **_PreserveStyles_**, **_PlainTextWordMail_**, **_ApplyFirstIndent_**, **_MatchParentheses_**, **_ReplaceDbDashes_**, **_ReplaceAutoSpaces_**|
| **wdDialogToolsOptionsAutoFormatAsYouType**| **_cmntrInWordMail_** , **_ApplyStylesHeadings_**, **_ApplyStylesHeadings_**, **_ApplyBorders_**, **_ApplyTables_**, **_ApplyDates_**, **_ApplyBulletedLists_**, **_ApplyNumberedLists_**, **_ApplyFirstIndent_**, **_ApplyClosings_**, **_ReplaceQuotes_**, **_ReplaceOrdinals_**, **_ReplaceFractions_**, **_ReplaceSymbols_**, **_ReplacePlainTextEmphasis_**, **_ReplaceHyperlinks_**, **_MatchParentheses_**, **_ReplaceAutoSpaces_**, **_ReplaceDbDashes_**, **_FormatListItemBeginning_**, **_TabIndent_**, **_DefineStyles_**, **_InsertOvers_**, **_InsertClosings_**, **_AutoLetterWizard_**, **_ShowOptionsFor_**, **_ApplyStylesLists_**, **_ApplySkipList_**, **_ApplyStylesOtherParas_**, **_ReplaceBullets_**, **_AdjustParaMarks_**, **_AdjustTabsSpaces_**, **_AdjustEmptyParas_**, **_PreserveStyles_**|
| **wdDialogToolsOptionsBidi**| **_DocViewDir_** , **_AddCtrlCopy_**, **_HebDoubleQuote_**, **_Numbers_**, **_Move_**, **_Sel_**, **_BiDirectional_**, **_ShowDiac_**, **_DiffDiacColor_**, **_Date_**, **_AdvanceHijri_**, **_MasterDocDir_**, **_OutlineDir_**, **_DiacriticColorVal_**, **_SequenceCheck_**, **_TypeNReplace_**|
| **wdDialogToolsOptionsCompatibility**| **_Product_** , **_Default_**, **_NoTabHangIndent_**, **_NoSpaceRaiseLower_**, **_PrintColBlack_**, **_WrapTrailSpaces_**, **_NoColumnBalance_**, **_ConvMailMergeEsc_**, **_SuppressSpBfAfterPgBrk_**, **_SuppressTopSpacing_**, **_OrigWordTableRules_**, **_TransparentMetafiles_**, **_ShowBreaksInFrames_**, **_SwapBordersFacingPages_**, **_LeaveBackslashAlone_**, **_ExpandShiftReturn_**, **_DontULTrailSpace_**, **_DontBalanceSbDbWidth_**, **_SuppressTopSpacingMac5_**, **_SpacingInWholePoints_**, **_PrintBodyTextBeforeHeader_**, **_NoLeading_**, **_NoSpaceForUL_**, **_MWSmallCaps_**, **_NoExtraLineSpacing_**, **_TruncateFontHeight_**, **_SubFontBySize_**, **_UsePrinterMetrics_**, **_WW6BorderRules_**, **_ExactOnTop_**, **_SuppressBottomSpacing_**, **_WPSpaceWidth_**, **_WPJustification_**, **_LineWrapLikeWord6_**, **_SpLayoutLikeWW8_**, **_FtnLayoutLikeWW8_**, **_DontUseHTMLParagraphAutoSpacing_**, **_DontAdjustLineHeightInTable_**, **_ForgetLastTabAlignment_**, **_UseAutospaceForFullWidthAlpha_**, **_AlignTablesRowByRow_**, **_LayoutRawTableWidth_**, **_LayoutTableRowsApart_**, **_UseWord97LineBreakingRules_**, **_DontBreakWrappedTables_**, **_DontSnapToGridInCell_**, **_DontAllowFieldEndSelect_**, **_ApplyBreakingRules_**, **_DontWrapTextWithPunct_**, **_DontUseAsianBreakRules_**, **_UseWord2002TableStyleRules_**, **_GrowAutofit_**|
| **wdDialogToolsOptionsEdit**| **_ReplaceSelection_** , **_DragAndDrop_**, **_AutoWordSelection_**, **_InsForPaste_**, **_Overtype_**, **_SmartCursoring_**, **_SmartCutPaste_**, **_AllowAccentedUppercase_**, **_PictureEditor_**, **_TabIndent_**, **_BsParaAlign_**, **_InlineConversion_**, **_IMELosingFocus_**, **_AllowClickAndTypeMouse_**, **_ClickAndTypeParagraphStyle_**, **_AutoKeyBi_**, **_PictureWrapType_**, **_SmartParaSelection_**, **_HypCtrlClickFollow_**, **_PasteRecovery_**, **_PromptUpdateStyle_**, **_FormatScanning_**, **_ShowFormatError_**|
| **wdDialogToolsOptionsEditCopyPaste**| **_SmartSentenceWordSpacing_** , **_SmartParaPaste_**, **_SmartTablePaste_**, **_SmartStylePaste_**, **_FormatPowerpointPaste_**, **_FormatExcelPaste_**, **_PasteMergeLists_**, **_CopyPasteDefaultOptions_**|
| **wdDialogToolsOptionsFileLocations**| **_Path_** , **_Setting_**|
| **wdDialogToolsOptionsFuzzy**| **_FuzzyCase_** , **_FuzzyByte_**, **_FuzzyHira_**, **_FuzzySmKana_**, **_FuzzyMinus_**, **_FuzzyRepSymbol_**, **_FuzzyKanji_**, **_FuzzyOldKana_**, **_FuzzyLongVowel_**, **_FuzzyDZ_**, **_FuzzyBV_**, **_FuzzyTC_**, **_FuzzyHF_**, **_FuzzyZJ_**, **_FuzzyAY_**, **_FuzzyKIKU_**, **_FuzzyPunct_**, **_FuzzySpace_**|
| **wdDialogToolsOptionsGeneral**| **_Pagination_** , **_WPHelp_**, **_WPDocNavKeys_**, **_BlueScreen_**, **_ErrorBeeps_**, **_Effects3d_**, **_UpdateLinks_**, **_SendMailAttach_**, **_RecentFiles_**, **_RecentFileCount_**, **_Units_**, **_ButtonFieldClicks_**, **_ShortMenuNames_**, **_RTFInClipboard_**, **_ConfirmConversions_**, **_TipWizardActive_**, **_AnimatedCursors_**, **_VirusProtection_**, **_SeparateFont_**, **_InterpretHIANSIToDBC_**, **_ExitWithRestoreSession_**, **_AsianText_**, **_PixelsInDialogs_**, **_UseCharacterUnit_**, **_BackgroundOpen_**, **_AutoCreateNewDrawings_**, **_AllowReadingMode_**|
| **wdDialogToolsOptionsPrint**| **_Draft_** , **_Reverse_**, **_UpdateFields_**, **_Summary_**, **_ShowCodes_**, **_Annotations_**, **_ShowHidden_**, **_EnvFeederInstalled_**, **_WidowControl_**, **_DfltTrueType_**, **_UpdateLinks_**, **_Background_**, **_DrawingObjects_**, **_FormsData_**, **_DefaultTray_**, **_PSOverText_**, **_MapPaperSize_**, **_FractionalWidths_**, **_PrOrder1_**, **_PrOrder2_**, **_PrintXmlTags_**, **_Backgrounds_**|
| **wdDialogToolsOptionsSave**| **_CreateBackup_** , **_FastSaves_**, **_SummaryPrompt_**, **_GlobalDotPrompt_**, **_NativePictureFormat_**, **_EmbedFonts_**, **_FormsData_**, **_AutoSave_**, **_SaveInterval_**, **_Password_**, **_WritePassword_**, **_RecommendReadOnly_**, **_SubsetFonts_**, **_BackgroundSave_**, **_DefaultSaveFormat_**, **_AddCtrlSave_**, **_DoNotEmbed_**, **_LocalNetworkFile_**, **_WordCompatibilityList_**, **_EmbedSmartTags_**, **_SmartTagXML_**, **_EmbedLinguisticData_**|
| **wdDialogToolsOptionsSecurity**| **_WarnMarkup_** , **_StoreRsid_**, **_ShowMarkupOpenSave_**|
| **wdDialogToolsOptionsSmartTag**| **_LabelSmartTags_** , **_ShowSmartTagOOUI_**|
| **wdDialogToolsOptionsSpellingAndGrammar**| **_AlwaysSuggest_** , **_SuggestFromMainDictOnly_**, **_IgnoreAllCaps_**, **_IgnoreMixedDigits_**, **_ResetIgnoreAll_**, **_Type_**, **_CustomDict1_**, **_CustomDict2_**, **_CustomDict3_**, **_CustomDict4_**, **_CustomDict5_**, **_CustomDict6_**, **_CustomDict7_**, **_CustomDict8_**, **_CustomDict9_**, **_CustomDict10_**, **_AutomaticSpellChecking_**, **_FilenamesEmailAliases_**, **_UserDict1_**, **_AutomaticGrammarChecking_**, **_ForegroundGrammar_**, **_ShowStatistics_**, **_Options_**, **_RecheckDocument_**, **_IgnoreAuxFind_**, **_IgnoreMissDictSearch_**, **_HideGrammarErrors_**, **_CheckSpelling_**, **_GrLidUI_**, **_SpLidUI_**, **_DictLang1_**, **_DictLang2_**, **_DictLang3_**, **_DictLang4_**, **_DictLang5_**, **_DictLang6_**, **_DictLang7_**, **_DictLang8_**, **_DictLang9_**, **_DictLang10_**, **_HideSpellingErrors_**, **_HebSpellStart_**, **_InitialAlefHamza_**, **_FinalYaa_**, **_GermanPostReformSpell_**, **_AraSpeller_**, **_ProcessCompoundNoun_**|
| **wdDialogToolsOptionsTrackChanges**| **_InsertedTextMark_** , **_InsertedTextColor_**, **_DeletedTextMark_**, **_DeletedTextColor_**, **_RevisedLinesMark_**, **_RevisedLinesColor_**, **_HighlightColor_**, **_RevisedPropertiesMark_**, **_RevisedPropertiesColor_**|
| **wdDialogToolsOptionsTypography**| **_KerningPairs_** , **_Justification_**, **_PunctLevel_**, **_FollowingPunct_**, **_LeadingPunct_**, **_ApplyToTemplate_**, **_JapaneseKinsokuStrict_**, **_FarEastLineBreakLanguage_**|
| **wdDialogToolsOptionsUserInfo**| **_Name_** , **_Initials_**, **_Address_**|
| **wdDialogToolsOptionsView**| **_DraftFont_** , **_WrapToWindow_**, **_PicturePlaceHolders_**, **_FieldCodes_**, **_BookMarks_**, **_FieldShading_**, **_StatusBar_**, **_HScroll_**, **_VScroll_**, **_StyleAreaWidth_**, **_Tabs_**, **_Spaces_**, **_Paras_**, **_Hyphens_**, **_Hidden_**, **_ShowAll_**, **_Drawings_**, **_Anchors_**, **_TextBoundaries_**, **_VRuler_**, **_Highlight_**, **_ShowAnimation_**, **_ScrnTp_**, **_LeftScroll_**, **_RRuler_**, **_OptionalBreak_**, **_EnlargeFontsLessThan_**, **_BrowseToWindow_**, **_PageBoundaries_**, **_WindowsInTaskbar_**, **_SmartTags_**, **_ShowAtStartup_**, **_Backgrounds_**|
| **wdDialogToolsProtectDocument**| **_DocumentPassword_** , **_NoReset_**, **_Type_**, **_UseDRM_**|
| **wdDialogToolsProtectSection**| **_Protect_** , **_Section_**|
| **wdDialogToolsRevisions**| **_MarkRevisions_** , **_ViewRevisions_**, **_PrintRevisions_**, **_AcceptAll_**, **_RejectAll_**|
| **wdDialogToolsSpellingAndGrammar**| **_SuggestionListBox_** , **_ForegroundGrammar_**|
| **wdDialogToolsTemplates**| **_Store_** , **_Template_**, **_LinkStyles_**|
| **wdDialogToolsThesaurus**|(none)|
| **wdDialogToolsUnprotectDocument**| **_DocumentPassword_**|
| **wdDialogToolsWordCount**| **_CountFootnotes_** , **_Pages_**, **_Words_**, **_Characters_**, **_DBCs_**, **_SBCs_**, **_CharactersIncludingSpaces_**, **_Paragraphs_**, **_Lines_**|
| **wdDialogTwoLinesInOne**|(none)|
| **wdDialogUpdateTOC**|(none)|
| **wdDialogViewZoom**| **_AutoFit_** , **_TwoPages_**, **_FullPage_**, **_NumColumns_**, **_NumRows_**, **_ZoomPercent_**, **_TextFit_**|
| **wdDialogWebOptions**|(none)|
| **wdDialogWindowActivate**| **_Window_**|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
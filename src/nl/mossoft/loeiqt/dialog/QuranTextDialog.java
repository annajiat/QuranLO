/*
 * This file is part of QuranLO
 *
 * Copyright (C) 2020 <mossie@mossoft.nl>
 *
 * QuranLO is free software: you can redistribute it and/or modify it under the terms of the GNU
 * General Public License as published by the Free Software Foundation, either version 3 of the
 * License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
 * even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License along with this program. If
 * not, see <https://www.gnu.org/licenses/>.
 */

package nl.mossoft.loeiqt.dialog;

import com.sun.star.awt.ActionEvent;
import com.sun.star.awt.ItemEvent;
import com.sun.star.awt.SpinEvent;
import com.sun.star.awt.XActionListener;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XItemListener;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XNumericField;
import com.sun.star.awt.XSpinField;
import com.sun.star.awt.XSpinListener;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameContainer;
import com.sun.star.frame.XController;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.style.ParagraphAdjust;
import com.sun.star.text.ControlCharacter;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import java.awt.Font;
import java.awt.GraphicsEnvironment;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import nl.mossoft.loeiqt.helper.DocumentHelper;
import nl.mossoft.loeiqt.helper.QuranReader;

/**
 * Creates the Dialog for the Quran parameters.
 * 
 * @author abdullah
 *
 */
public class QuranTextDialog {

  private static final char A = '\u0061'; //
  private static final char ALIF = '\u0627'; // ALIF
  private static final short ALIGNMENT_RIGHT = (short) 2;
  private static final String LPAR = "\uFD3E"; // Arabic Left parentheses
  private static final String RPAR = "\uFD3F"; // Arabic Right parentheses

  private static final String AYATCHKBXID = "AyatAllCheckBox";
  private static final String AYATCHKBXLABELID = "AyatCheckBoxLabel";
  private static final String AYATCHKBXLABELTXT = "All Ayat:";
  private static final String AYATFROMLABELID = "AyatFromLabel";
  private static final String AYATFROMLABELTXT = "From:";
  private static final String AYATFROMNUMFLDID = "AyatFromNumericField";
  private static final String AYATGRPBXID = "AyatGroupBox";
  private static final String AYATGRPBXTXT = "Select Ayat";
  private static final String AYATTOLABELID = "AyatToLabel";
  private static final String AYATTOLABELTTXT = "To:";
  private static final String AYATTONUMFLDID = "AyatToNumericField";

  private static final String FONTARABICFONTSIZEID = "ArabicFontSize";
  private static final String FONTARABICFONTSIZELABELTXT = "Fontsize:";
  private static final String FONTARABICFONTSIZENUMFLDID = "ArabicFontSizeNumericField";
  private static final String FONTARABICGRPBXID = "FontArabicGroupBox";
  private static final String FONTARABICGRPBXTXT = "Arabic Font";
  private static final String FONTARABICLABELID = "ArabicFont";
  private static final String FONTARABICLABELTXT = "Font:";
  private static final String FONTARABICLSTBXID = "FontArabicListBox";
  private static final String FONTARABICNMBRBTTNID = "FontArabicNumberButton";
  private static final String FONTARABICNMBRBTTNLABEL = "\uFD3E" + "\u06F1" + "\uFD3F";
  private static final String FONTARABICBOLDBTTNID = "FontArabicBoldButton";
  private static final String FONTARABICBOLDBTTNLABEL = "B";
  private static final String FONTNONARABICFONTSIZEID = "NonArabicFontSize";
  private static final String FONTNONARABICFONTSIZELABELTXT = "Fontsize:";
  private static final String FONTNONARABICFONTSIZENUMFLDID = "NonArabicFontSizeNumericField";
  private static final String FONTNONARABICGRPBXID = "FontNonArabicGroupBox";
  private static final String FONTNONARABICGRPBXTXT = "Non-Arabic Font";
  private static final String FONTNONARABICLABELID = "NonArabicFont";
  private static final String FONTNONARABICLABELTXT = "Font:";
  private static final String FONTNONARABICLSTBXID = "FontNonArabicListBox";

  private static final String LANGARABICCHKBXID = "LangArabicCheckBox";
  private static final String LANGARABICLABELID = "LangArabicLabel";
  private static final String LANGARABICLABELTXT = "Arabic:";
  private static final String LANGARABICLSTBXID = "LangArabicListBox";
  private static final String LANGGRPBXID = "LangGroupBox";
  private static final String LANGGRPBXLABEL = "Languages";
  private static final String LANGTRNSLTNCHKBXID = "LangTranslationCheckBox";
  private static final String LANGTRNSLTNLABELID = "LangTranslationLabel";
  private static final String LANGTRNSLTNLABELTXT = "Translation:";
  private static final String LANGTRNSLTNLSTBXID = "LangTranslationListBox";
  private static final String LANGTRNSLTRTNCHKBXID = "LangTransliterationCheckBox";
  private static final String LANGTRNSLTRTNLABELID = "LangTransliterationLabel";
  private static final String LANGTRNSLTRTNLABELTXT = "Transliteration:";
  private static final String LANGTRNSLTRTNLSTBXID = "LangTransliterationListBox";

  private static final String MISCGRPBXID = "MiscGroupBox";
  private static final String MISCGRPBXLABEL = "Miscellaneous";
  private static final String MISCLBLCHKBXID = "MiscLineByLineCheckBox:";
  private static final String MISCLBLCHKBXLABELID = "MiscLineByLineLabel:";
  private static final String MISCLBLCHKBXLABELTXT = "Line by Line:";

  private static final String OKBTTNID = "OkButton";
  private static final String OKBTTNTXT = "Ok";

  private static final String SURAHGRPBXID = "SurahGroupBox";
  private static final String SURAHGRPBXLABEL = "Select Surah";
  private static final String SURAHGRPBXLABELID = "SurahName:";
  private static final String SURAHGRPBXLABELTXT = "Name:";
  private static final String SURAHLSTBXID = "SurahListBox";

  private static String getItemLanguague(String item) {
    final String[] itemsSelected = item.split("[(]");
    return itemsSelected[0].trim();
  }

  private static String getItemVersion(String item) {
    final String[] itemsSelected = item.split("[(]");
    return itemsSelected[1].replace(")", " ").trim().replace(" ", "_");
  }

  public static QuranTextDialog newInstance(XComponentContext context,
      XMultiComponentFactory factory) {
    return new QuranTextDialog(context, factory);
  }

  private static boolean shortToBoolean(short s) {
    return s != 0;
  }

  private double defaultArabicCharHeight;
  private String defaultArabicFontName;
  private double defaultNonArabicCharHeight;
  private String defaultNonArabicFontName;

  private XCheckBox dlgArabicCheckBox;
  private XListBox dlgArabicFontListBox;
  private XNumericField dlgArabicFontsizeNumericField;
  private XSpinField dlgArabicFontsizeNumericSpinfield;
  private XListBox dlgArabicListBox;
  private XButton dlgArabicNumberButton;
  private XButton dlgArabicBoldButton;
  private XSpinField dlgAyaFromNumericSpinfield;
  private XCheckBox dlgAyatAllCheckBox;
  private XNumericField dlgAyatFromNumericField;
  private XSpinField dlgAyaToNumericSpinfield;
  private XNumericField dlgAyatToNumericField;
  private final XComponentContext dlgComponentContext;
  private XControlContainer dlgControlContainer;
  private XDialog dlgDialog;
  private XCheckBox dlgLbLCheckBox;
  private final XMultiComponentFactory dlgMultiComponentFactory;
  private XMultiServiceFactory dlgMFactory;
  private XNameContainer dlgNameContainer;
  private XListBox dlgNonArabicFontListBox;
  private XNumericField dlgNonArabicFontsizeNumericField;
  private XSpinField dlgNonArabicFontsizeNumericSpinfield;
  private XButton dlgOkButton;
  private XListBox dlgSurahListBox;
  private XCheckBox dlgTrnsltnCheckBox;
  private XListBox dlgTrnsltnListBox;
  private XCheckBox dlgTrnsltrtnCheckBox;
  private XListBox dlgTrnsltrtnListBox;

  private boolean selectedAllRngInd = true;
  private String selectedArbcFntNm;
  private double selectedArbcFntSz;
  private boolean selectedArbcInd = false;
  private String selectedArbcLngg;
  private String selectedArbcVrsn;
  private long selectedAytFrm = 1;
  private long selectedAytT = 1;
  private boolean selectedLbLInd = false;
  private final boolean selectedLnNmbrInd = true;
  private String selectedNnArbcFntNm;
  private double selectedNnArbcFntSz;
  private int selectedSurhNo = 1;
  private String selectedTrnsltnLngg;
  private String selectedTrnsltnVrsn;
  private boolean selectedTrnsltrtnInd = false;
  private String selectedTrnsltrtnLngg;
  private String selectedTrnsltrtnVrsn;
  private boolean selTrnsltnInd = false;

  private QuranTextDialog(XComponentContext context, XMultiComponentFactory factory) {

    this.dlgComponentContext = context;
    this.dlgMultiComponentFactory = factory;

    try {
      this.getLoDocumentDefaults();
      this.dlgDialog = this.createDialog();
    } catch (final Exception e) {
      e.printStackTrace();
    }
  }

  private void addParagraph(XText text, XParagraphCursor paragraphCursor, String paragraph,
      ParagraphAdjust alignment, short writingMode)
      throws com.sun.star.lang.IllegalArgumentException, UnknownPropertyException,
      PropertyVetoException, WrappedTargetException {

    paragraphCursor.gotoEndOfParagraph(false);
    text.insertControlCharacter(paragraphCursor, ControlCharacter.PARAGRAPH_BREAK, false);

    final XPropertySet paragraphCursorPropertySet = DocumentHelper.getPropertySet(paragraphCursor);
    paragraphCursorPropertySet.setPropertyValue("ParaAdjust", alignment);
    paragraphCursorPropertySet.setPropertyValue("WritingMode", writingMode);
    text.insertString(paragraphCursor, paragraph, false);
  }

  private XDialog createDialog() throws Exception {
    final Object dlgModel = this.dlgMultiComponentFactory.createInstanceWithContext(
        "com.sun.star.awt.UnoControlDialogModel", this.dlgComponentContext);

    this.dlgMFactory = UnoRuntime.queryInterface(XMultiServiceFactory.class, dlgModel);
    this.dlgNameContainer = UnoRuntime.queryInterface(XNameContainer.class, dlgModel);

    final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, dlgModel);
    pset.setPropertyValue("PositionX", Integer.valueOf(100));
    pset.setPropertyValue("PositionY", Integer.valueOf(100));
    pset.setPropertyValue("Width", Integer.valueOf(295));
    pset.setPropertyValue("Height", Integer.valueOf(200));
    pset.setPropertyValue("Title", "Insert Quran Text");

    // create the dialog control and set the model
    final Object dialog = this.dlgMultiComponentFactory
        .createInstanceWithContext("com.sun.star.awt.UnoControlDialog", this.dlgComponentContext);

    final XControl xControl = UnoRuntime.queryInterface(XControl.class, dialog);

    this.dlgControlContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);

    final XControlModel xControlModel = UnoRuntime.queryInterface(XControlModel.class, dlgModel);
    xControl.setModel(xControlModel);

    final int X = 5;
    final int Y = 5;

    // Surah Group
    this.insertGroupBox(QuranTextDialog.SURAHGRPBXID, QuranTextDialog.SURAHGRPBXLABEL, X, Y, 30,
        140, true);
    this.insertLabel(QuranTextDialog.SURAHGRPBXLABELID, QuranTextDialog.SURAHGRPBXLABELTXT,
        QuranTextDialog.ALIGNMENT_RIGHT, X + 5, Y + 12 + 1, 10, 45, true);
    this.insertSurahListBox(QuranTextDialog.SURAHLSTBXID, X + 68, Y + 12, 10, 70, true, 0);

    // Ayat Group
    this.insertGroupBox(QuranTextDialog.AYATGRPBXID, QuranTextDialog.AYATGRPBXTXT, X, Y + 30, 70,
        140, true);
    this.insertLabel(QuranTextDialog.AYATCHKBXLABELID, QuranTextDialog.AYATCHKBXLABELTXT,
        QuranTextDialog.ALIGNMENT_RIGHT, X + 5, Y + 42 + 1, 10, 45, true);
    this.insertLabel(QuranTextDialog.AYATFROMLABELID, QuranTextDialog.AYATFROMLABELTXT,
        QuranTextDialog.ALIGNMENT_RIGHT, X + 5, Y + 62 + 1, 10, 45, false);
    this.insertLabel(QuranTextDialog.AYATTOLABELID, QuranTextDialog.AYATTOLABELTTXT,
        QuranTextDialog.ALIGNMENT_RIGHT, X + 5, Y + 82 + 1, 10, 45, false);
    this.insertAyatAllCheckBox(QuranTextDialog.AYATCHKBXID, X + 54, Y + 42 - 1, 10, 10, true, 1);
    this.insertAyatFromNumericField(QuranTextDialog.AYATFROMNUMFLDID, X + 68, Y + 62, 10, 70, false,
        3);
    this.insertAyatToNumericField(QuranTextDialog.AYATTONUMFLDID, X + 68, Y + 82, 10, 70, false, 3);

    // Languages
    this.insertGroupBox(QuranTextDialog.LANGGRPBXID, QuranTextDialog.LANGGRPBXLABEL, X, Y + 100, 70,
        140, true);
    this.insertLabel(QuranTextDialog.LANGARABICLABELID, QuranTextDialog.LANGARABICLABELTXT,
        QuranTextDialog.ALIGNMENT_RIGHT, X + 5, Y + 112 + 1, 10, 45, true);
    this.insertArabicCheckBox(QuranTextDialog.LANGARABICCHKBXID, X + 54, Y + 112 - 1, 10, 10, true,
        1);
    this.insertArabicListBox(QuranTextDialog.LANGARABICLSTBXID, X + 68, Y + 112 - 1, 10, 70, true,
        1);

    this.insertLabel(QuranTextDialog.LANGTRNSLTNLABELID, QuranTextDialog.LANGTRNSLTNLABELTXT,
        QuranTextDialog.ALIGNMENT_RIGHT, X + 5, Y + 132 + 1, 10, 45, true);
    this.insertTranslationCheckBox(QuranTextDialog.LANGTRNSLTNCHKBXID, X + 54, Y + 132 - 1, 10, 10,
        true, 1);
    this.insertTranslationListBox(QuranTextDialog.LANGTRNSLTNLSTBXID, X + 68, Y + 132 - 1, 10, 70,
        false, 1);

    // TODO needs to be set to true when transliteration is ready.
    // insertLabel(LANGTRNSLTRTNLABELID, LANGTRNSLTRTNLABELTXT, ALIGNMENT_RIGHT, X + 5, Y + 152 + 1,
    // 10, 45, false);
    // TODO needs to be set to true when transliteration is ready.
    // insertTransliterationCheckBox(LANGTRNSLTRTNCHKBXID, X + 54, Y + 152 - 1, 10, 10, false, 0);
    // insertTransliterationListBox(LANGTRNSLTRTNLSTBXID, X + 68, Y + 152 - 1, 10, 70, false, 1);

    // Arabic font
    this.insertGroupBox(QuranTextDialog.FONTARABICGRPBXID, QuranTextDialog.FONTARABICGRPBXTXT,
        X + 145, Y, 70, 140, true);
    this.insertLabel(QuranTextDialog.FONTARABICLABELID, QuranTextDialog.FONTARABICLABELTXT,
        QuranTextDialog.ALIGNMENT_RIGHT, X + 150, Y + 12 + 1, 10, 45, true);
    this.insertArabicFontListBox(QuranTextDialog.FONTARABICLSTBXID, X + 145 + 68, Y + 12, 10, 70,
        true, 1);
    this.insertLabel(QuranTextDialog.FONTARABICFONTSIZEID,
        QuranTextDialog.FONTARABICFONTSIZELABELTXT, QuranTextDialog.ALIGNMENT_RIGHT, X + 150,
        Y + 32 + 1, 10, 45, true);
    this.insertArabicFontSizeNumericField(QuranTextDialog.FONTARABICFONTSIZENUMFLDID, X + 145 + 68,
        Y + 32, 10, 70, true, 3);
    // insertArabicBoldButton(FONTARABICBOLDBTTNID, FONTARABICBOLDBTTNLABEL, X + 145 + 68, Y + 49,
    // 13,
    // 13, true, 3);
    // insertArabicNumberButton(FONTARABICNMBRBTTNID, FONTARABICNMBRBTTNLABEL, X + 145 + 17 + 68,
    // Y + 49, 13, 13, true, 3);

    // Non-Arabic font
    this.insertGroupBox(QuranTextDialog.FONTNONARABICGRPBXID, QuranTextDialog.FONTNONARABICGRPBXTXT,
        X + 145, Y + 70, 70, 140, false);
    this.insertLabel(QuranTextDialog.FONTNONARABICLABELID, QuranTextDialog.FONTNONARABICLABELTXT,
        QuranTextDialog.ALIGNMENT_RIGHT, X + 150, Y + 75 + 12 + 1, 10, 45, false);
    this.insertNonArabicFontListBox(QuranTextDialog.FONTNONARABICLSTBXID, X + 145 + 68, Y + 70 + 12,
        10, 70, false, 1);
    this.insertLabel(QuranTextDialog.FONTNONARABICFONTSIZEID,
        QuranTextDialog.FONTNONARABICFONTSIZELABELTXT, QuranTextDialog.ALIGNMENT_RIGHT, X + 150,
        Y + 75 + 32 + 1, 10, 45, false);
    this.insertNonArabicFontSizeNumericField(QuranTextDialog.FONTNONARABICFONTSIZENUMFLDID,
        X + 145 + 68, Y + 70 + 32, 10, 70, false, 3);

    this.insertGroupBox(QuranTextDialog.MISCGRPBXID, QuranTextDialog.MISCGRPBXLABEL, X + 145,
        Y + 140, 30, 140, true);
    this.insertLabel(QuranTextDialog.MISCLBLCHKBXLABELID, QuranTextDialog.MISCLBLCHKBXLABELTXT,
        QuranTextDialog.ALIGNMENT_RIGHT, X + 150, Y + 140 + 12 + 1, 10, 45, true);
    this.insertLineByLineCheckBox(QuranTextDialog.MISCLBLCHKBXID, X + 145 + 68, Y + 140 + 12 - 1,
        10, 10, true, 1);

    // Cancel & Ok Buttons
    this.insertOkButton(QuranTextDialog.OKBTTNID, QuranTextDialog.OKBTTNTXT, X + 245, Y + 170 + 5,
        12, 40, true, 1);


    // Create a peer
    final Object toolkit = this.dlgMultiComponentFactory
        .createInstanceWithContext("com.sun.star.awt.ExtToolkit", this.dlgComponentContext);
    final XToolkit xToolkit = UnoRuntime.queryInterface(XToolkit.class, toolkit);
    final XWindow xWindow = UnoRuntime.queryInterface(XWindow.class, xControl);
    xWindow.setVisible(false);
    xControl.createPeer(xToolkit, null);

    return UnoRuntime.queryInterface(XDialog.class, dialog);
  }

  public void enableComponent(String componentId, boolean enabled) {
    final XControl xControl =
        UnoRuntime.queryInterface(XControl.class, this.dlgControlContainer.getControl(componentId));
    final XPropertySet xPropertySet =
        UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
    try {
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));
    } catch (IllegalArgumentException | UnknownPropertyException | PropertyVetoException
        | WrappedTargetException e) {
      e.printStackTrace();
    }
  }

  void execute() {
    this.dlgDialog.execute();
  }

  private String getAyahLine(int surahno, int ayahno, String language, String version) {

    final QuranReader qr = new QuranReader(language, version, this.dlgComponentContext);
    String line = ((ayahno == 1) && (surahno != 1 && surahno != 9))
        ? this.getBismillah(surahno, language, version) + " "
            + qr.getAyahNoOfSuraNo(surahno, ayahno)
        : qr.getAyahNoOfSuraNo(surahno, ayahno);

    if (this.selectedLnNmbrInd) {
      if (language.equals("Arabic")) {
        line = line + " " + RPAR + QuranReader.numToArabNum(ayahno, selectedArbcFntNm) + LPAR + " ";
      } else {
        line = "(" + ayahno + ") " + line;
      }
    }
    return line;
  }

  private String getBismillah(int surahno, String language, String version) {
    final QuranReader qr = new QuranReader(language, version, this.dlgComponentContext);
    return qr.getBismillah();
  }

  private double getDefaultArabicCharHeight() {
    return this.defaultArabicCharHeight;
  }

  private String getDefaultArabicFontName() {
    return this.defaultArabicFontName;
  }

  private double getDefaultNonArabicCharHeight() {
    return this.defaultNonArabicCharHeight;
  }

  private String getDefaultNonArabicFontName() {
    return this.defaultNonArabicFontName;
  }

  private void getLoDocumentDefaults() {
    final XTextDocument textDoc = DocumentHelper.getCurrentDocument(this.dlgComponentContext);
    final XController controller = textDoc.getCurrentController();
    final XTextViewCursorSupplier textViewCursorSupplier =
        DocumentHelper.getCursorSupplier(controller);
    final XTextViewCursor textViewCursor = textViewCursorSupplier.getViewCursor();
    final XText text = textViewCursor.getText();
    final XTextCursor textCursor = text.createTextCursorByRange(textViewCursor.getStart());
    final XParagraphCursor paragraphCursor =
        UnoRuntime.queryInterface(XParagraphCursor.class, textCursor);
    final XPropertySet paragraphCursorPropertySet = DocumentHelper.getPropertySet(paragraphCursor);

    try {
      this.defaultArabicCharHeight =
          (float) paragraphCursorPropertySet.getPropertyValue("CharHeightComplex");
    } catch (UnknownPropertyException | WrappedTargetException e) {
      this.defaultArabicCharHeight = 10;
    }
    try {
      this.defaultArabicFontName =
          (String) paragraphCursorPropertySet.getPropertyValue("CharFontNameComplex");
    } catch (UnknownPropertyException | WrappedTargetException e) {
      this.defaultArabicFontName = "No Default set";
    }
    try {
      this.defaultNonArabicFontName =
          (String) paragraphCursorPropertySet.getPropertyValue("CharFontName");
    } catch (UnknownPropertyException | WrappedTargetException e) {
      this.defaultNonArabicFontName = "No Default set";
    }
    try {
      this.defaultNonArabicCharHeight =
          (float) paragraphCursorPropertySet.getPropertyValue("CharHeight");
    } catch (UnknownPropertyException | WrappedTargetException e) {
      this.defaultNonArabicCharHeight = 10;
    }

  }

  public int getSurahNo() {
    return this.selectedSurhNo;
  }

  /**
   * @param componentID
   * @param label
   * @param posx
   * @param posy
   * @param height
   * @param width
   * @param enabled
   * @param tabIndex
   */
  public void insertArabicBoldButton(String componentID, String label, int posx, int posy,
      int height, int width, boolean enabled, int tabIndex) {
    try {
      final Object ButtonModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlButtonModel");

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, ButtonModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Label", label);
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("Toggle", Boolean.valueOf(true));
      pset.setPropertyValue("State", Short.valueOf((short) 1));
      pset.setPropertyValue("Align", Short.valueOf((short) 1));
      pset.setPropertyValue("FontRelief", com.sun.star.text.FontRelief.EMBOSSED);

      this.dlgNameContainer.insertByName(componentID, ButtonModel);
      this.dlgArabicBoldButton = UnoRuntime.queryInterface(XButton.class,
          this.dlgControlContainer.getControl(componentID));


      this.dlgArabicBoldButton.addActionListener(new XActionListener() {

        @Override
        public void actionPerformed(ActionEvent arg0) {

        }

        @Override
        public void disposing(EventObject arg0) {}

      });

    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }


  /**
   * @param componentID
   * @param posx
   * @param posy
   * @param height
   * @param width
   * @param enabled
   * @param tabIndex
   */
  public void insertArabicCheckBox(String componentID, int posx, int posy, int height, int width,
      boolean enabled, int tabIndex) {
    try {
      final Object checkBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, checkBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("State", Short.valueOf((short) 1));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("State", Short.valueOf((short) 1));

      this.selectedArbcInd = QuranTextDialog.shortToBoolean((short) 1);
      this.dlgNameContainer.insertByName(componentID, checkBoxModel);

      this.dlgArabicCheckBox = UnoRuntime.queryInterface(XCheckBox.class,
          this.dlgControlContainer.getControl(componentID));
      this.dlgArabicCheckBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          QuranTextDialog.this.selectedArbcInd =
              QuranTextDialog.shortToBoolean(QuranTextDialog.this.dlgArabicCheckBox.getState());

          QuranTextDialog.this.enableComponent(QuranTextDialog.LANGARABICLSTBXID,
              QuranTextDialog.this.selectedArbcInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTARABICGRPBXID,
              QuranTextDialog.this.selectedArbcInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTARABICLABELID,
              QuranTextDialog.this.selectedArbcInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTARABICLSTBXID,
              QuranTextDialog.this.selectedArbcInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTARABICFONTSIZEID,
              QuranTextDialog.this.selectedArbcInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTARABICFONTSIZENUMFLDID,
              QuranTextDialog.this.selectedArbcInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.OKBTTNID,
              QuranTextDialog.this.selectedArbcInd || QuranTextDialog.this.selTrnsltnInd
                  || QuranTextDialog.this.selectedTrnsltrtnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.MISCGRPBXID,
              QuranTextDialog.this.selectedArbcInd || QuranTextDialog.this.selTrnsltnInd
                  || QuranTextDialog.this.selectedTrnsltrtnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.MISCLBLCHKBXID,
              QuranTextDialog.this.selectedArbcInd || QuranTextDialog.this.selTrnsltnInd
                  || QuranTextDialog.this.selectedTrnsltrtnInd);
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertArabicFontListBox(String componentID, int posx, int posy, int height, int width,
      boolean enabled, int tabIndex) {
    try {
      final Object ListBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      this.selectedArbcFntNm = this.getDefaultArabicFontName();

      final List<String> list = new ArrayList<String>();
      final Locale locale = new Locale.Builder().setScript("ARAB").build();

      final short[] selectedItems = new short[1];

      final String[] fonts =
          GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames(locale);
      for (final String font : fonts) {
        if (new Font(font, Font.PLAIN, 10).canDisplay(QuranTextDialog.ALIF)) {
          list.add(font);
          if (font.equals(this.selectedArbcFntNm)) {
            selectedItems[0] = (short) (list.size() - 1);
          }
        }
      }
      final String[] itemList = list.toArray(new String[0]);

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));
      pset.setPropertyValue("StringItemList", itemList);
      pset.setPropertyValue("SelectedItems", selectedItems);

      pset.setPropertyValue("MultiSelection", Boolean.FALSE);
      pset.setPropertyValue("Dropdown", Boolean.TRUE);

      this.dlgNameContainer.insertByName(componentID, ListBoxModel);

      this.dlgArabicFontListBox = UnoRuntime.queryInterface(XListBox.class,
          this.dlgControlContainer.getControl(componentID));

      this.dlgArabicFontListBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {

        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertArabicFontSizeNumericField(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      final Object NumericFieldModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlNumericFieldModel");

      this.selectedArbcFntSz = this.getDefaultArabicCharHeight();

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, NumericFieldModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("Spin", Boolean.TRUE);
      pset.setPropertyValue("DecimalAccuracy", Short.valueOf((short) 1));

      this.dlgNameContainer.insertByName(componentID, NumericFieldModel);

      this.dlgArabicFontsizeNumericField = UnoRuntime.queryInterface(XNumericField.class,
          this.dlgControlContainer.getControl(componentID));

      this.dlgArabicFontsizeNumericField.setValue(this.defaultArabicCharHeight);
      this.dlgArabicFontsizeNumericField.setMin(1);

      this.dlgArabicFontsizeNumericSpinfield = UnoRuntime.queryInterface(XSpinField.class,
          this.dlgControlContainer.getControl(componentID));
      this.dlgArabicFontsizeNumericSpinfield.addSpinListener(new XSpinListener() {
        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void down(SpinEvent arg0) {
          QuranTextDialog.this.selectedArbcFntSz =
              QuranTextDialog.this.dlgArabicFontsizeNumericField.getValue();
        }

        @Override
        public void first(SpinEvent arg0) {}

        @Override
        public void last(SpinEvent arg0) {}

        @Override
        public void up(SpinEvent arg0) {
          QuranTextDialog.this.selectedArbcFntSz =
              QuranTextDialog.this.dlgArabicFontsizeNumericField.getValue();
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertArabicListBox(String componentID, int posx, int posy, int height, int width,
      boolean enabled, int tabIndex) {
    try {
      final Object ListBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      final List<String> list = new ArrayList<String>();

      QuranReader.getAllQuranVersions().entrySet().stream().forEach((entry) -> {
        final String currentKey = entry.getKey();
        final String[] currentValue = entry.getValue();
        for (final String element : currentValue) {
          if (currentKey.equals("Arabic")) {
            list.add(currentKey + " (" + element + ")");
          }
        }
      });

      final String[] itemList = list.toArray(new String[0]);

      final short[] selectedItems = new short[] {(short) 0};

      this.selectedArbcLngg = QuranTextDialog.getItemLanguague(itemList[0]);
      this.selectedArbcVrsn = QuranTextDialog.getItemVersion(itemList[0]);

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));
      pset.setPropertyValue("StringItemList", itemList);
      pset.setPropertyValue("SelectedItems", selectedItems);

      pset.setPropertyValue("MultiSelection", Boolean.FALSE);
      pset.setPropertyValue("Dropdown", Boolean.TRUE);

      this.dlgNameContainer.insertByName(componentID, ListBoxModel);

      this.dlgArabicListBox = UnoRuntime.queryInterface(XListBox.class,
          this.dlgControlContainer.getControl(componentID));

      this.dlgArabicListBox.addItemListener(new XItemListener() {
        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          QuranTextDialog.this.selectedArbcLngg = QuranTextDialog
              .getItemLanguague(QuranTextDialog.this.dlgArabicListBox.getSelectedItem());
          QuranTextDialog.this.selectedArbcVrsn = QuranTextDialog
              .getItemVersion(QuranTextDialog.this.dlgArabicListBox.getSelectedItem());
        }

      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertArabicNumberButton(String componentID, String label, int posx, int posy,
      int height, int width, boolean enabled, int tabIndex) {
    try {
      final Object ButtonModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlButtonModel");

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, ButtonModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Label", label);
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("Toggle", Boolean.valueOf(true));
      pset.setPropertyValue("State", Short.valueOf((short) 1));
      pset.setPropertyValue("Align", Short.valueOf((short) 1));
      pset.setPropertyValue("FontRelief", com.sun.star.text.FontRelief.EMBOSSED);

      this.dlgNameContainer.insertByName(componentID, ButtonModel);
      this.dlgArabicNumberButton = UnoRuntime.queryInterface(XButton.class,
          this.dlgControlContainer.getControl(componentID));


      this.dlgArabicNumberButton.addActionListener(new XActionListener() {

        @Override
        public void actionPerformed(ActionEvent arg0) {

        }

        @Override
        public void disposing(EventObject arg0) {}

      });

    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertAyatAllCheckBox(String componentID, int posx, int posy, int height, int width,
      boolean enabled, int tabIndex) {
    try {
      final Object checkBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, checkBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("State", Short.valueOf((short) 1));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("State", Short.valueOf((short) 1));

      this.selectedAllRngInd = QuranTextDialog.shortToBoolean((short) 1);
      this.dlgNameContainer.insertByName(componentID, checkBoxModel);

      this.dlgAyatAllCheckBox = UnoRuntime.queryInterface(XCheckBox.class,
          this.dlgControlContainer.getControl(componentID));
      this.dlgAyatAllCheckBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          QuranTextDialog.this.selectedAllRngInd =
              QuranTextDialog.shortToBoolean(QuranTextDialog.this.dlgAyatAllCheckBox.getState());

          if (!QuranTextDialog.this.selectedAllRngInd) {
            QuranTextDialog.this.dlgAyatToNumericField.setValue(QuranTextDialog.this.selectedAytT);
            QuranTextDialog.this.dlgAyatFromNumericField
                .setValue(QuranTextDialog.this.selectedAytFrm);
          }

          QuranTextDialog.this.enableComponent(QuranTextDialog.AYATFROMLABELID,
              !QuranTextDialog.this.selectedAllRngInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.AYATFROMNUMFLDID,
              !QuranTextDialog.this.selectedAllRngInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.AYATTOLABELID,
              !QuranTextDialog.this.selectedAllRngInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.AYATTONUMFLDID,
              !QuranTextDialog.this.selectedAllRngInd);
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertAyatFromNumericField(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      final Object NumericFieldModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlNumericFieldModel");

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, NumericFieldModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("Spin", Boolean.TRUE);
      pset.setPropertyValue("DecimalAccuracy", Short.valueOf((short) 0));

      this.dlgNameContainer.insertByName(componentID, NumericFieldModel);

      this.dlgAyatFromNumericField = UnoRuntime.queryInterface(XNumericField.class,
          this.dlgControlContainer.getControl(componentID));

      this.dlgAyatFromNumericField.setValue(this.selectedAytFrm);
      this.dlgAyatFromNumericField.setMin(this.selectedAytFrm);

      this.dlgAyaFromNumericSpinfield = UnoRuntime.queryInterface(XSpinField.class,
          this.dlgControlContainer.getControl(componentID));
      this.dlgAyaFromNumericSpinfield.addSpinListener(new XSpinListener() {
        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void down(SpinEvent arg0) {
          QuranTextDialog.this.selectedAytFrm =
              Math.round(QuranTextDialog.this.dlgAyatFromNumericField.getValue());
        }

        @Override
        public void first(SpinEvent arg0) {}

        @Override
        public void last(SpinEvent arg0) {}

        @Override
        public void up(SpinEvent arg0) {
          QuranTextDialog.this.selectedAytFrm =
              Math.round(QuranTextDialog.this.dlgAyatFromNumericField.getValue());
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertAyatToNumericField(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      final Object NumericFieldModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlNumericFieldModel");


      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, NumericFieldModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("Spin", Boolean.TRUE);
      pset.setPropertyValue("DecimalAccuracy", Short.valueOf((short) 0));

      this.dlgNameContainer.insertByName(componentID, NumericFieldModel);

      this.dlgAyatToNumericField = UnoRuntime.queryInterface(XNumericField.class,
          this.dlgControlContainer.getControl(componentID));

      this.dlgAyatToNumericField.setValue(this.selectedAytT);
      this.dlgAyatToNumericField.setMax(this.selectedAytT);

      this.dlgAyaToNumericSpinfield = UnoRuntime.queryInterface(XSpinField.class,
          this.dlgControlContainer.getControl(componentID));
      this.dlgAyaToNumericSpinfield.addSpinListener(new XSpinListener() {
        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void down(SpinEvent arg0) {
          QuranTextDialog.this.selectedAytT =
              Math.round(QuranTextDialog.this.dlgAyatToNumericField.getValue());
        }

        @Override
        public void first(SpinEvent arg0) {}

        @Override
        public void last(SpinEvent arg0) {}

        @Override
        public void up(SpinEvent arg0) {
          QuranTextDialog.this.selectedAytT =
              Math.round(QuranTextDialog.this.dlgAyatToNumericField.getValue());
          System.out.printf("Selected Ayat To: %s\n", QuranTextDialog.this.selectedAytT);
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  /**
   * Insert a group box in the dialog.
   * 
   * @param componentID identification of the group box.
   * @param label label of the group box.
   * @param posx x position.
   * @param posy y position.
   * @param height height of group box.
   * @param width width of group box.
   * @param enabled true if group box is enabled.
   */
  public void insertGroupBox(String componentID, String label, int posx, int posy, int height,
      int width, boolean enabled) {
    try {
      final Object groupBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlGroupBoxModel");

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, groupBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("Label", label);
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      this.dlgNameContainer.insertByName(componentID, groupBoxModel);
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  /**
   * Insert a label in the dialog.
   * 
   * @param componentID identification of the group box.
   * @param label label of the group box.
   * @param posx x position.
   * @param posy y position.
   * @param height height of group box.
   * @param width width of group box.
   * @param enabled true if group box is enabled.
   */
  public void insertLabel(String componentID, String label, short alignment, int posx, int posy,
      int height, int width, boolean enabled) {
    try {
      final Object fixedTextModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlFixedTextModel");

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, fixedTextModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("Label", label);
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("Align", alignment);

      this.dlgNameContainer.insertByName(componentID, fixedTextModel);
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  /**
   * Insert check box for indicating the format in the dialog.
   * 
   * @param componentID identification of the check box.
   * @param posx x position.
   * @param posy y position.
   * @param height height of check box.
   * @param width width of check box.
   * @param enabled true if the Line by line format is selected.
   */
  public void insertLineByLineCheckBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      final Object checkBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, checkBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("State", Short.valueOf((short) 1));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("State", Short.valueOf((short) 1));

      this.selectedLbLInd = QuranTextDialog.shortToBoolean((short) 1);
      this.dlgNameContainer.insertByName(componentID, checkBoxModel);

      this.dlgLbLCheckBox = UnoRuntime.queryInterface(XCheckBox.class,
          this.dlgControlContainer.getControl(componentID));
      this.dlgLbLCheckBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          QuranTextDialog.this.selectedLbLInd =
              QuranTextDialog.shortToBoolean(QuranTextDialog.this.dlgLbLCheckBox.getState());
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  /**
   * Insert list box for selecting the non-arabic font in the dialog.
   * 
   * @param componentID identification of the list box.
   * @param posx x position.
   * @param posy y position.
   * @param height height of list box.
   * @param width width of list box.
   * @param enabled true if the list box is enabled.
   */
  public void insertNonArabicFontListBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      final Object ListBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      this.selectedNnArbcFntNm = this.getDefaultNonArabicFontName();

      final List<String> list = new ArrayList<String>();
      final Locale locale = new Locale.Builder().setScript("LATN").build();

      final short[] selectedItems = new short[1];

      final String[] fonts =
          GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames(locale);
      for (final String font : fonts) {
        if (new Font(font, Font.PLAIN, 10).canDisplay(QuranTextDialog.A)) {
          list.add(font);
          if (font.equals(this.selectedNnArbcFntNm)) {
            selectedItems[0] = (short) (list.size() - 1);
          }
        }
      }

      final String[] itemList = list.toArray(new String[0]);

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));
      pset.setPropertyValue("StringItemList", itemList);
      pset.setPropertyValue("SelectedItems", selectedItems);

      pset.setPropertyValue("MultiSelection", Boolean.FALSE);
      pset.setPropertyValue("Dropdown", Boolean.TRUE);

      this.dlgNameContainer.insertByName(componentID, ListBoxModel);

      this.dlgNonArabicFontListBox = UnoRuntime.queryInterface(XListBox.class,
          this.dlgControlContainer.getControl(componentID));

      this.dlgNonArabicFontListBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {

        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertNonArabicFontSizeNumericField(String componentID, int posx, int posy,
      int height, int width, boolean enabled, int tabIndex) {
    try {
      final Object NumericFieldModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlNumericFieldModel");

      this.selectedNnArbcFntSz = this.getDefaultNonArabicCharHeight();

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, NumericFieldModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("Spin", Boolean.TRUE);
      pset.setPropertyValue("DecimalAccuracy", Short.valueOf((short) 1));

      this.dlgNameContainer.insertByName(componentID, NumericFieldModel);

      this.dlgNonArabicFontsizeNumericField = UnoRuntime.queryInterface(XNumericField.class,
          this.dlgControlContainer.getControl(componentID));

      this.dlgNonArabicFontsizeNumericField.setValue(this.defaultNonArabicCharHeight);
      this.dlgNonArabicFontsizeNumericField.setMin(1);

      this.dlgNonArabicFontsizeNumericSpinfield = UnoRuntime.queryInterface(XSpinField.class,
          this.dlgControlContainer.getControl(componentID));
      this.dlgNonArabicFontsizeNumericSpinfield.addSpinListener(new XSpinListener() {
        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void down(SpinEvent arg0) {
          QuranTextDialog.this.selectedNnArbcFntSz =
              QuranTextDialog.this.dlgNonArabicFontsizeNumericField.getValue();
        }

        @Override
        public void first(SpinEvent arg0) {}

        @Override
        public void last(SpinEvent arg0) {}

        @Override
        public void up(SpinEvent arg0) {
          QuranTextDialog.this.selectedNnArbcFntSz =
              QuranTextDialog.this.dlgNonArabicFontsizeNumericField.getValue();
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertOkButton(String componentID, String label, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      final Object ButtonModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlButtonModel");

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, ButtonModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Label", label);
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      this.dlgNameContainer.insertByName(componentID, ButtonModel);
      this.dlgOkButton = UnoRuntime.queryInterface(XButton.class,
          this.dlgControlContainer.getControl(componentID));


      this.dlgOkButton.addActionListener(new XActionListener() {

        @Override
        public void actionPerformed(ActionEvent arg0) {
          QuranTextDialog.this.writeSurah(QuranTextDialog.this.selectedSurhNo);
          QuranTextDialog.this.dlgDialog.endExecute();
        }

        @Override
        public void disposing(EventObject arg0) {}

      });

    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertSurahListBox(String componentID, int posx, int posy, int height, int width,
      boolean enabled, int tabIndex) {
    try {
      final Object ListBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      final String[] itemList = new String[114];
      for (int i = 0; i < 114; i++) {
        itemList[i] = QuranReader.getSurahName(i) + " (" + (i + 1) + ")";
      }
      final short[] selectedItems = new short[] {(short) 0};

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("MultiSelection", Boolean.FALSE);
      pset.setPropertyValue("Dropdown", Boolean.TRUE);
      pset.setPropertyValue("StringItemList", itemList);
      pset.setPropertyValue("SelectedItems", selectedItems);

      this.selectedSurhNo = 1;
      this.selectedAytFrm = 1;
      this.selectedAytT = QuranReader.getSurahSize(this.selectedSurhNo - 1);

      this.dlgNameContainer.insertByName(componentID, ListBoxModel);

      this.dlgSurahListBox = UnoRuntime.queryInterface(XListBox.class,
          this.dlgControlContainer.getControl(componentID));
      this.dlgSurahListBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          QuranTextDialog.this.selectedSurhNo =
              (short) (QuranTextDialog.this.dlgSurahListBox.getSelectedItemPos() + 1);

          QuranTextDialog.this.selectedAytFrm = 1;
          QuranTextDialog.this.dlgAyatFromNumericField
              .setValue(QuranTextDialog.this.selectedAytFrm);
          QuranTextDialog.this.dlgAyatFromNumericField.setMin(QuranTextDialog.this.selectedAytFrm);

          QuranTextDialog.this.selectedAytT =
              QuranReader.getSurahSize(QuranTextDialog.this.selectedSurhNo - 1);
          QuranTextDialog.this.dlgAyatToNumericField.setValue(QuranTextDialog.this.selectedAytT);
          QuranTextDialog.this.dlgAyatToNumericField.setMax(QuranTextDialog.this.selectedAytT);
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertTranslationCheckBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      final Object checkBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, checkBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("State", Short.valueOf((short) 1));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("State", Short.valueOf((short) 0));
      this.selTrnsltnInd = QuranTextDialog.shortToBoolean(Short.valueOf((short) 0));

      this.dlgNameContainer.insertByName(componentID, checkBoxModel);

      this.dlgTrnsltnCheckBox = UnoRuntime.queryInterface(XCheckBox.class,
          this.dlgControlContainer.getControl(componentID));
      this.dlgTrnsltnCheckBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {

          QuranTextDialog.this.selTrnsltnInd =
              QuranTextDialog.shortToBoolean(QuranTextDialog.this.dlgTrnsltnCheckBox.getState());

          QuranTextDialog.this.enableComponent(QuranTextDialog.LANGTRNSLTNLSTBXID,
              QuranTextDialog.this.selTrnsltnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTNONARABICGRPBXID,
              QuranTextDialog.this.selectedTrnsltrtnInd || QuranTextDialog.this.selTrnsltnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTNONARABICLABELID,
              QuranTextDialog.this.selectedTrnsltrtnInd || QuranTextDialog.this.selTrnsltnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTNONARABICLSTBXID,
              QuranTextDialog.this.selectedTrnsltrtnInd || QuranTextDialog.this.selTrnsltnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTNONARABICFONTSIZEID,
              QuranTextDialog.this.selectedTrnsltrtnInd || QuranTextDialog.this.selTrnsltnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTNONARABICFONTSIZENUMFLDID,
              QuranTextDialog.this.selectedTrnsltrtnInd || QuranTextDialog.this.selTrnsltnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.OKBTTNID,
              QuranTextDialog.this.selectedArbcInd || QuranTextDialog.this.selTrnsltnInd
                  || QuranTextDialog.this.selectedTrnsltrtnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.MISCGRPBXID,
              QuranTextDialog.this.selectedArbcInd || QuranTextDialog.this.selTrnsltnInd
                  || QuranTextDialog.this.selectedTrnsltrtnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.MISCLBLCHKBXID,
              QuranTextDialog.this.selectedArbcInd || QuranTextDialog.this.selTrnsltnInd
                  || QuranTextDialog.this.selectedTrnsltrtnInd);
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertTranslationListBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      final Object ListBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      final List<String> list = new ArrayList<String>();

      QuranReader.getAllQuranVersions().entrySet().stream().forEach((entry) -> {
        final String currentKey = entry.getKey();
        final String[] currentValue = entry.getValue();
        for (final String element : currentValue) {
          if (!currentKey.equals("Arabic")) {
            list.add(currentKey + " (" + element + ")");
          }
        }
      });

      final String[] itemList = list.toArray(new String[0]);

      final short[] selectedItems = new short[] {(short) 0};

      this.selectedTrnsltnLngg = QuranTextDialog.getItemLanguague(itemList[0]);
      this.selectedTrnsltnVrsn = QuranTextDialog.getItemVersion(itemList[0]);

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));
      pset.setPropertyValue("StringItemList", itemList);
      pset.setPropertyValue("SelectedItems", selectedItems);

      pset.setPropertyValue("MultiSelection", Boolean.FALSE);
      pset.setPropertyValue("Dropdown", Boolean.TRUE);

      this.dlgNameContainer.insertByName(componentID, ListBoxModel);

      this.dlgTrnsltnListBox = UnoRuntime.queryInterface(XListBox.class,
          this.dlgControlContainer.getControl(componentID));

      this.dlgTrnsltnListBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          QuranTextDialog.this.selectedTrnsltnLngg = QuranTextDialog
              .getItemLanguague(QuranTextDialog.this.dlgTrnsltnListBox.getSelectedItem());
          QuranTextDialog.this.selectedTrnsltnVrsn = QuranTextDialog
              .getItemVersion(QuranTextDialog.this.dlgTrnsltnListBox.getSelectedItem());
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertTransliterationCheckBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      final Object checkBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, checkBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("State", Short.valueOf((short) 1));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      pset.setPropertyValue("State", Short.valueOf((short) 0));

      this.selectedAllRngInd = QuranTextDialog.shortToBoolean((short) 1);
      this.dlgNameContainer.insertByName(componentID, checkBoxModel);

      this.dlgTrnsltrtnCheckBox = UnoRuntime.queryInterface(XCheckBox.class,
          this.dlgControlContainer.getControl(componentID));
      this.dlgTrnsltrtnCheckBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          QuranTextDialog.this.selectedTrnsltrtnInd =
              QuranTextDialog.shortToBoolean(QuranTextDialog.this.dlgTrnsltrtnCheckBox.getState());

          QuranTextDialog.this.enableComponent(QuranTextDialog.LANGTRNSLTRTNLSTBXID,
              QuranTextDialog.this.selectedTrnsltrtnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTNONARABICGRPBXID,
              QuranTextDialog.this.selectedTrnsltrtnInd || QuranTextDialog.this.selTrnsltnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTNONARABICLABELID,
              QuranTextDialog.this.selectedTrnsltrtnInd || QuranTextDialog.this.selTrnsltnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTNONARABICLSTBXID,
              QuranTextDialog.this.selectedTrnsltrtnInd || QuranTextDialog.this.selTrnsltnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTNONARABICFONTSIZEID,
              QuranTextDialog.this.selectedTrnsltrtnInd || QuranTextDialog.this.selTrnsltnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.FONTNONARABICFONTSIZENUMFLDID,
              QuranTextDialog.this.selectedTrnsltrtnInd || QuranTextDialog.this.selTrnsltnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.OKBTTNID,
              QuranTextDialog.this.selectedArbcInd || QuranTextDialog.this.selTrnsltnInd
                  || QuranTextDialog.this.selectedTrnsltrtnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.MISCGRPBXID,
              QuranTextDialog.this.selectedArbcInd || QuranTextDialog.this.selTrnsltnInd
                  || QuranTextDialog.this.selectedTrnsltrtnInd);
          QuranTextDialog.this.enableComponent(QuranTextDialog.MISCLBLCHKBXID,
              QuranTextDialog.this.selectedArbcInd || QuranTextDialog.this.selTrnsltnInd
                  || QuranTextDialog.this.selectedTrnsltrtnInd);
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertTransliterationListBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      final Object ListBoxModel =
          this.dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      final List<String> list = new ArrayList<String>();

      QuranReader.getAllQuranVersions().entrySet().stream().forEach((entry) -> {
        final String currentKey = entry.getKey();
        final String[] currentValue = entry.getValue();
        for (final String element : currentValue) {
          if (currentKey.equals("Transliteration")) {
            list.add(currentKey + " (" + element + ")");
          }
        }
      });

      final String[] itemList = list.toArray(new String[0]);

      final short[] selectedItems = new short[] {(short) 0};

      this.selectedTrnsltrtnLngg = QuranTextDialog.getItemLanguague(itemList[0]);
      this.selectedTrnsltrtnVrsn = QuranTextDialog.getItemVersion(itemList[0]);

      final XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      pset.setPropertyValue("PositionX", Integer.valueOf(posx));
      pset.setPropertyValue("PositionY", Integer.valueOf(posy));
      pset.setPropertyValue("Width", Integer.valueOf(width));
      pset.setPropertyValue("Height", Integer.valueOf(height));
      pset.setPropertyValue("Name", componentID);
      pset.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      pset.setPropertyValue("Enabled", Boolean.valueOf(enabled));
      pset.setPropertyValue("StringItemList", itemList);
      pset.setPropertyValue("SelectedItems", selectedItems);

      pset.setPropertyValue("MultiSelection", Boolean.FALSE);
      pset.setPropertyValue("Dropdown", Boolean.TRUE);

      this.dlgNameContainer.insertByName(componentID, ListBoxModel);

      this.dlgTrnsltrtnListBox = UnoRuntime.queryInterface(XListBox.class,
          this.dlgControlContainer.getControl(componentID));

      this.dlgTrnsltrtnListBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          QuranTextDialog.this.selectedTrnsltrtnLngg = QuranTextDialog
              .getItemLanguague(QuranTextDialog.this.dlgTrnsltrtnListBox.getSelectedItem());
          QuranTextDialog.this.selectedTrnsltrtnVrsn = QuranTextDialog
              .getItemVersion(QuranTextDialog.this.dlgTrnsltrtnListBox.getSelectedItem());
        }
      });
    } catch (final com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void writeSurah(int surahNumber) {
    final XTextDocument textDoc = DocumentHelper.getCurrentDocument(this.dlgComponentContext);
    final XController controller = textDoc.getCurrentController();
    final XTextViewCursorSupplier textViewCursorSupplier =
        DocumentHelper.getCursorSupplier(controller);
    final XTextViewCursor textViewCursor = textViewCursorSupplier.getViewCursor();
    final XText text = textViewCursor.getText();
    final XTextCursor textCursor = text.createTextCursorByRange(textViewCursor.getStart());
    final XParagraphCursor paragraphCursor =
        UnoRuntime.queryInterface(XParagraphCursor.class, textCursor);
    final XPropertySet paragraphCursorPropertySet = DocumentHelper.getPropertySet(paragraphCursor);

    try {
      paragraphCursorPropertySet.setPropertyValue("CharFontName", this.selectedNnArbcFntNm);
      paragraphCursorPropertySet.setPropertyValue("CharFontNameComplex", this.selectedArbcFntNm);
      paragraphCursorPropertySet.setPropertyValue("CharHeight", this.selectedNnArbcFntSz);
      paragraphCursorPropertySet.setPropertyValue("CharHeightComplex", this.selectedArbcFntSz);

      final int from = (this.selectedAllRngInd) ? 1 : (int) this.selectedAytFrm;
      final int to = (this.selectedAllRngInd) ? QuranReader.getSurahSize(surahNumber - 1) + 1
          : (int) this.selectedAytT + 1;

      if (this.selectedLbLInd) {
        String linea = " ";
        String lineb = "";
        String linec = "";;
        for (int l = from; l < to; l++) {
          if (this.selectedArbcInd) {
            linea = linea
                + this.getAyahLine(surahNumber, l, this.selectedArbcLngg, this.selectedArbcVrsn)
                + "\n";
          }
          if (this.selTrnsltnInd) {
            lineb = lineb + this.getAyahLine(surahNumber, l, this.selectedTrnsltnLngg,
                this.selectedTrnsltnVrsn) + "\n";
          }
          if (this.selectedTrnsltrtnInd) {
            linec = linec + this.getAyahLine(surahNumber, l, this.selectedTrnsltrtnLngg,
                this.selectedTrnsltrtnVrsn) + "\n";
          }
        }
        this.addParagraph(text, paragraphCursor, linea + "\n",
            com.sun.star.style.ParagraphAdjust.RIGHT, com.sun.star.text.WritingMode2.RL_TB);
        this.addParagraph(text, paragraphCursor, lineb + "\n",
            com.sun.star.style.ParagraphAdjust.LEFT, com.sun.star.text.WritingMode2.LR_TB);
        this.addParagraph(text, paragraphCursor, linec + "\n",
            com.sun.star.style.ParagraphAdjust.LEFT, com.sun.star.text.WritingMode2.LR_TB);
      } else {
        if (this.selectedArbcInd) {
          String linea = "";
          for (int l = from; l < to; l++) {
            linea = linea
                + this.getAyahLine(surahNumber, l, this.selectedArbcLngg, this.selectedArbcVrsn)
                + " ";
          }
          this.addParagraph(text, paragraphCursor, linea + "\n",
              com.sun.star.style.ParagraphAdjust.RIGHT, com.sun.star.text.WritingMode2.RL_TB);
        }
        if (this.selTrnsltnInd) {
          String lineb = "";
          for (int l = from; l < to; l++) {
            lineb = lineb + this.getAyahLine(surahNumber, l, this.selectedTrnsltnLngg,
                this.selectedTrnsltnVrsn) + " ";
          }
          this.addParagraph(text, paragraphCursor, lineb + "\n",
              com.sun.star.style.ParagraphAdjust.LEFT, com.sun.star.text.WritingMode2.LR_TB);
        }
        if (this.selectedTrnsltrtnInd) {
          String linec = "";
          for (int l = from; l < to; l++) {
            linec = linec + this.getAyahLine(surahNumber, l, this.selectedTrnsltrtnLngg,
                this.selectedTrnsltrtnVrsn) + " ";
          }
          this.addParagraph(text, paragraphCursor, linec + "\n",
              com.sun.star.style.ParagraphAdjust.LEFT, com.sun.star.text.WritingMode2.LR_TB);
        }
      }
    } catch (com.sun.star.lang.IllegalArgumentException | UnknownPropertyException
        | PropertyVetoException | WrappedTargetException e) {
      e.printStackTrace();
    }
  }

}

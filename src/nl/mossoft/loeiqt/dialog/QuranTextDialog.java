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
 * @author abdullah
 *
 */
public class QuranTextDialog {

  private static final char A = '\u0061'; //
  private static final char ALIF = '\u0627'; // ALIF
  private static final short ALIGNMENT_RIGHT = (short) 2;

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
    String[] itemsSelected = item.split("[(]");
    return itemsSelected[0].trim();
  }

  private static String getItemVersion(String item) {
    String[] itemsSelected = item.split("[(]");
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
  private XComponentContext dlgComponentContext;
  private XControlContainer dlgControlContainer;
  private XDialog dlgDialog;
  private XCheckBox dlgLbLCheckBox;
  private XMultiComponentFactory dlgMultiComponentFactory;
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
  private boolean selectedLnNmbrInd = true;
  private String selectedNnArbcFntNm;
  private double selectedNnArbcFntSz;
  private int selectedSurhNo = 1;
  private String selectedTrnsltnLngg;
  private String selectedTrnsltnVrsn;
  private boolean selectedTrnsltrtnInd = false;
  private String selectedTrnsltrtnLngg;
  private String selectedTrnsltrtnVrsn;
  private boolean selTrnsltnInd = false;

  public int getSurahNo() {
    return selectedSurhNo;
  }

  private QuranTextDialog(XComponentContext context, XMultiComponentFactory factory) {

    dlgComponentContext = context;
    dlgMultiComponentFactory = factory;

    try {
      getLoDocumentDefaults();
      dlgDialog = createDialog();
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private void addParagraph(XText text, XParagraphCursor paragraphCursor, String paragraph,
      ParagraphAdjust alignment, short writingMode)
      throws com.sun.star.lang.IllegalArgumentException, UnknownPropertyException,
      PropertyVetoException, WrappedTargetException {

    paragraphCursor.gotoEndOfParagraph(false);
    text.insertControlCharacter(paragraphCursor, ControlCharacter.PARAGRAPH_BREAK, false);

    XPropertySet paragraphCursorPropertySet = DocumentHelper.getPropertySet(paragraphCursor);
    paragraphCursorPropertySet.setPropertyValue("ParaAdjust", alignment);
    paragraphCursorPropertySet.setPropertyValue("WritingMode", writingMode);
    text.insertString(paragraphCursor, paragraph, false);
  }

  private XDialog createDialog() throws Exception {
    Object dlgModel = dlgMultiComponentFactory
        .createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", dlgComponentContext);

    dlgMFactory = UnoRuntime.queryInterface(XMultiServiceFactory.class, dlgModel);
    dlgNameContainer = UnoRuntime.queryInterface(XNameContainer.class, dlgModel);

    XPropertySet pset = UnoRuntime.queryInterface(XPropertySet.class, dlgModel);
    pset.setPropertyValue("PositionX", Integer.valueOf(100));
    pset.setPropertyValue("PositionY", Integer.valueOf(100));
    pset.setPropertyValue("Width", Integer.valueOf(295));
    pset.setPropertyValue("Height", Integer.valueOf(200));
    pset.setPropertyValue("Title", "Insert Quran Text");

    // create the dialog control and set the model
    Object dialog = dlgMultiComponentFactory
        .createInstanceWithContext("com.sun.star.awt.UnoControlDialog", dlgComponentContext);

    XControl xControl = UnoRuntime.queryInterface(XControl.class, dialog);

    dlgControlContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);

    XControlModel xControlModel = UnoRuntime.queryInterface(XControlModel.class, dlgModel);
    xControl.setModel(xControlModel);

    final int X = 5;
    final int Y = 5;

    // Surah Group
    insertGroupBox(SURAHGRPBXID, SURAHGRPBXLABEL, X, Y, 30, 140, true);
    insertLabel(SURAHGRPBXLABELID, SURAHGRPBXLABELTXT, ALIGNMENT_RIGHT, X + 5, Y + 12 + 1, 10, 45,
        true);
    insertSurahListBox(SURAHLSTBXID, X + 68, Y + 12, 10, 70, true, 0);

    // Ayat Group
    insertGroupBox(AYATGRPBXID, AYATGRPBXTXT, X, Y + 30, 70, 140, true);
    insertLabel(AYATCHKBXLABELID, AYATCHKBXLABELTXT, ALIGNMENT_RIGHT, X + 5, Y + 42 + 1, 10, 45,
        true);
    insertLabel(AYATFROMLABELID, AYATFROMLABELTXT, ALIGNMENT_RIGHT, X + 5, Y + 62 + 1, 10, 45,
        false);
    insertLabel(AYATTOLABELID, AYATTOLABELTTXT, ALIGNMENT_RIGHT, X + 5, Y + 82 + 1, 10, 45, false);
    insertAyatAllCheckBox(AYATCHKBXID, X + 54, Y + 42 - 1, 10, 10, true, 1);
    insertAyatFromNumericField(AYATFROMNUMFLDID, X + 68, Y + 62, 10, 70, false, 3);
    insertAyatToNumericField(AYATTONUMFLDID, X + 68, Y + 82, 10, 70, false, 3);

    // Languages
    insertGroupBox(LANGGRPBXID, LANGGRPBXLABEL, X, Y + 100, 70, 140, true);
    insertLabel(LANGARABICLABELID, LANGARABICLABELTXT, ALIGNMENT_RIGHT, X + 5, Y + 112 + 1, 10, 45,
        true);
    insertArabicCheckBox(LANGARABICCHKBXID, X + 54, Y + 112 - 1, 10, 10, true, 1);
    insertArabicListBox(LANGARABICLSTBXID, X + 68, Y + 112 - 1, 10, 70, true, 1);

    insertLabel(LANGTRNSLTNLABELID, LANGTRNSLTNLABELTXT, ALIGNMENT_RIGHT, X + 5, Y + 132 + 1, 10,
        45, true);
    insertTranslationCheckBox(LANGTRNSLTNCHKBXID, X + 54, Y + 132 - 1, 10, 10, true, 1);
    insertTranslationListBox(LANGTRNSLTNLSTBXID, X + 68, Y + 132 - 1, 10, 70, false, 1);

    // TODO needs to be set to true when transliteration is ready.
    // insertLabel(LANGTRNSLTRTNLABELID, LANGTRNSLTRTNLABELTXT, ALIGNMENT_RIGHT, X + 5, Y + 152 + 1,
    // 10, 45, false);
    // TODO needs to be set to true when transliteration is ready.
    // insertTransliterationCheckBox(LANGTRNSLTRTNCHKBXID, X + 54, Y + 152 - 1, 10, 10, false, 0);
    // insertTransliterationListBox(LANGTRNSLTRTNLSTBXID, X + 68, Y + 152 - 1, 10, 70, false, 1);

    // Arabic font
    insertGroupBox(FONTARABICGRPBXID, FONTARABICGRPBXTXT, X + 145, Y, 70, 140, true);
    insertLabel(FONTARABICLABELID, FONTARABICLABELTXT, ALIGNMENT_RIGHT, X + 150, Y + 12 + 1, 10, 45,
        true);
    insertArabicFontListBox(FONTARABICLSTBXID, X + 145 + 68, Y + 12, 10, 70, true, 1);
    insertLabel(FONTARABICFONTSIZEID, FONTARABICFONTSIZELABELTXT, ALIGNMENT_RIGHT, X + 150,
        Y + 32 + 1, 10, 45, true);
    insertArabicFontSizeNumericField(FONTARABICFONTSIZENUMFLDID, X + 145 + 68, Y + 32, 10, 70, true,
        3);
    // insertArabicBoldButton(FONTARABICBOLDBTTNID, FONTARABICBOLDBTTNLABEL, X + 145 + 68, Y + 49,
    // 13,
    // 13, true, 3);
    // insertArabicNumberButton(FONTARABICNMBRBTTNID, FONTARABICNMBRBTTNLABEL, X + 145 + 17 + 68,
    // Y + 49, 13, 13, true, 3);

    // Non-Arabic font
    insertGroupBox(FONTNONARABICGRPBXID, FONTNONARABICGRPBXTXT, X + 145, Y + 70, 70, 140, false);
    insertLabel(FONTNONARABICLABELID, FONTNONARABICLABELTXT, ALIGNMENT_RIGHT, X + 150,
        Y + 75 + 12 + 1, 10, 45, false);
    insertNonArabicFontListBox(FONTNONARABICLSTBXID, X + 145 + 68, Y + 70 + 12, 10, 70, false, 1);
    insertLabel(FONTNONARABICFONTSIZEID, FONTNONARABICFONTSIZELABELTXT, ALIGNMENT_RIGHT, X + 150,
        Y + 75 + 32 + 1, 10, 45, false);
    insertNonArabicFontSizeNumericField(FONTNONARABICFONTSIZENUMFLDID, X + 145 + 68, Y + 70 + 32,
        10, 70, false, 3);

    insertGroupBox(MISCGRPBXID, MISCGRPBXLABEL, X + 145, Y + 140, 30, 140, true);
    insertLabel(MISCLBLCHKBXLABELID, MISCLBLCHKBXLABELTXT, ALIGNMENT_RIGHT, X + 150,
        Y + 140 + 12 + 1, 10, 45, true);
    insertLineByLineCheckBox(MISCLBLCHKBXID, X + 145 + 68, Y + 140 + 12 - 1, 10, 10, true, 1);

    // Cancel & Ok Buttons
    insertOkButton(OKBTTNID, OKBTTNTXT, X + 245, Y + 170 + 5, 12, 40, true, 1);


    // Create a peer
    Object toolkit = dlgMultiComponentFactory
        .createInstanceWithContext("com.sun.star.awt.ExtToolkit", dlgComponentContext);
    XToolkit xToolkit = UnoRuntime.queryInterface(XToolkit.class, toolkit);
    XWindow xWindow = UnoRuntime.queryInterface(XWindow.class, xControl);
    xWindow.setVisible(false);
    xControl.createPeer(xToolkit, null);

    return UnoRuntime.queryInterface(XDialog.class, dialog);
  }

  public void enableComponent(String componentId, boolean enabled) {
    XControl xControl =
        UnoRuntime.queryInterface(XControl.class, dlgControlContainer.getControl(componentId));
    XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
    try {
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));
    } catch (IllegalArgumentException | UnknownPropertyException | PropertyVetoException
        | WrappedTargetException e) {
      e.printStackTrace();
    }
  }

  void execute() {
    dlgDialog.execute();
  }

  private String getAyahLine(int surahno, int ayahno, String language, String version) {
    final String LPAR = "\uFD3E"; // Arabic Left parentheses
    final String RPAR = "\uFD3F"; // Arabic Right parentheses

    QuranReader qr = new QuranReader(language, version, dlgComponentContext);
    String line = ((ayahno == 1) && (surahno != 1 && surahno != 9))
        ? getBismillah(surahno, language, version) + " " + qr.getAyahNoOfSuraNo(surahno, ayahno)
        : qr.getAyahNoOfSuraNo(surahno, ayahno);

    if (selectedLnNmbrInd) {
      if (language.equals("Arabic")) {
        line = line + " " + RPAR + QuranReader.numToArabNum(ayahno) + LPAR + " ";
      } else {
        line = "(" + ayahno + ") " + line;
      }
    }
    return line;
  }

  private String getBismillah(int surahno, String language, String version) {
    QuranReader qr = new QuranReader(language, version, dlgComponentContext);
    return qr.getBismillah();
  }

  private double getDefaultArabicCharHeight() {
    return defaultArabicCharHeight;
  }

  private String getDefaultArabicFontName() {
    return defaultArabicFontName;
  }

  private double getDefaultNonArabicCharHeight() {
    return defaultNonArabicCharHeight;
  }

  private String getDefaultNonArabicFontName() {
    return defaultNonArabicFontName;
  }

  private void getLoDocumentDefaults() {
    XTextDocument textDoc = DocumentHelper.getCurrentDocument(dlgComponentContext);
    XController controller = textDoc.getCurrentController();
    XTextViewCursorSupplier textViewCursorSupplier = DocumentHelper.getCursorSupplier(controller);
    XTextViewCursor textViewCursor = textViewCursorSupplier.getViewCursor();
    XText text = textViewCursor.getText();
    XTextCursor textCursor = text.createTextCursorByRange(textViewCursor.getStart());
    XParagraphCursor paragraphCursor =
        UnoRuntime.queryInterface(XParagraphCursor.class, textCursor);
    XPropertySet paragraphCursorPropertySet = DocumentHelper.getPropertySet(paragraphCursor);

    try {
      defaultArabicCharHeight =
          (float) paragraphCursorPropertySet.getPropertyValue("CharHeightComplex");
    } catch (UnknownPropertyException | WrappedTargetException e) {
      defaultArabicCharHeight = 10;
    }
    try {
      defaultArabicFontName =
          (String) paragraphCursorPropertySet.getPropertyValue("CharFontNameComplex");
    } catch (UnknownPropertyException | WrappedTargetException e) {
      defaultArabicFontName = "No Default set";
    }
    try {
      defaultNonArabicFontName =
          (String) paragraphCursorPropertySet.getPropertyValue("CharFontName");
    } catch (UnknownPropertyException | WrappedTargetException e) {
      defaultNonArabicFontName = "No Default set";
    }
    try {
      defaultNonArabicCharHeight =
          (float) paragraphCursorPropertySet.getPropertyValue("CharHeight");
    } catch (UnknownPropertyException | WrappedTargetException e) {
      defaultNonArabicCharHeight = 10;
    }

  }

  public void insertArabicCheckBox(String componentID, int posx, int posy, int height, int width,
      boolean enabled, int tabIndex) {
    try {
      Object checkBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, checkBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("State", Short.valueOf((short) 1));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("State", Short.valueOf((short) 1));

      selectedArbcInd = shortToBoolean((short) 1);
      dlgNameContainer.insertByName(componentID, checkBoxModel);

      dlgArabicCheckBox =
          UnoRuntime.queryInterface(XCheckBox.class, dlgControlContainer.getControl(componentID));
      dlgArabicCheckBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          selectedArbcInd = shortToBoolean(dlgArabicCheckBox.getState());

          enableComponent(LANGARABICLSTBXID, selectedArbcInd);
          enableComponent(FONTARABICGRPBXID, selectedArbcInd);
          enableComponent(FONTARABICLABELID, selectedArbcInd);
          enableComponent(FONTARABICLSTBXID, selectedArbcInd);
          enableComponent(FONTARABICFONTSIZEID, selectedArbcInd);
          enableComponent(FONTARABICFONTSIZENUMFLDID, selectedArbcInd);
          enableComponent(OKBTTNID, selectedArbcInd || selTrnsltnInd || selectedTrnsltrtnInd);
          enableComponent(MISCGRPBXID, selectedArbcInd || selTrnsltnInd || selectedTrnsltrtnInd);
          enableComponent(MISCLBLCHKBXID, selectedArbcInd || selTrnsltnInd || selectedTrnsltrtnInd);
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }


  public void insertArabicFontListBox(String componentID, int posx, int posy, int height, int width,
      boolean enabled, int tabIndex) {
    try {
      Object ListBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      selectedArbcFntNm = getDefaultArabicFontName();

      List<String> list = new ArrayList<String>();
      Locale locale = new Locale.Builder().setScript("ARAB").build();

      short[] selectedItems = new short[1];

      String[] fonts =
          GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames(locale);
      for (int i = 0; i < fonts.length; i++) {
        if (new Font(fonts[i], Font.PLAIN, 10).canDisplay(ALIF)) {
          list.add(fonts[i]);
          if (fonts[i].equals(selectedArbcFntNm)) {
            selectedItems[0] = (short) (list.size() - 1);
          }
        }
      }
      String[] itemList = list.toArray(new String[0]);

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));
      xPropertySet.setPropertyValue("StringItemList", itemList);
      xPropertySet.setPropertyValue("SelectedItems", selectedItems);

      xPropertySet.setPropertyValue("MultiSelection", Boolean.FALSE);
      xPropertySet.setPropertyValue("Dropdown", Boolean.TRUE);

      dlgNameContainer.insertByName(componentID, ListBoxModel);

      dlgArabicFontListBox =
          UnoRuntime.queryInterface(XListBox.class, dlgControlContainer.getControl(componentID));

      dlgArabicFontListBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {

        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertArabicFontSizeNumericField(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      Object NumericFieldModel =
          dlgMFactory.createInstance("com.sun.star.awt.UnoControlNumericFieldModel");

      selectedArbcFntSz = getDefaultArabicCharHeight();

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, NumericFieldModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("Spin", Boolean.TRUE);
      xPropertySet.setPropertyValue("DecimalAccuracy", Short.valueOf((short) 1));

      dlgNameContainer.insertByName(componentID, NumericFieldModel);

      dlgArabicFontsizeNumericField = UnoRuntime.queryInterface(XNumericField.class,
          dlgControlContainer.getControl(componentID));

      dlgArabicFontsizeNumericField.setValue(defaultArabicCharHeight);
      dlgArabicFontsizeNumericField.setMin(1);

      dlgArabicFontsizeNumericSpinfield =
          UnoRuntime.queryInterface(XSpinField.class, dlgControlContainer.getControl(componentID));
      dlgArabicFontsizeNumericSpinfield.addSpinListener(new XSpinListener() {
        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void down(SpinEvent arg0) {
          selectedArbcFntSz = dlgArabicFontsizeNumericField.getValue();
        }

        @Override
        public void first(SpinEvent arg0) {}

        @Override
        public void last(SpinEvent arg0) {}

        @Override
        public void up(SpinEvent arg0) {
          selectedArbcFntSz = dlgArabicFontsizeNumericField.getValue();
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertArabicListBox(String componentID, int posx, int posy, int height, int width,
      boolean enabled, int tabIndex) {
    try {
      Object ListBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      List<String> list = new ArrayList<String>();

      QuranReader.getAllQuranVersions().entrySet().stream().forEach((entry) -> {
        String currentKey = entry.getKey();
        String[] currentValue = entry.getValue();
        for (int i = 0; i < currentValue.length; i++) {
          if (currentKey.equals("Arabic")) {
            list.add(currentKey + " (" + currentValue[i] + ")");
          }
        }
      });

      String[] itemList = list.toArray(new String[0]);

      short[] selectedItems = new short[] {(short) 0};

      selectedArbcLngg = getItemLanguague(itemList[0]);
      selectedArbcVrsn = getItemVersion(itemList[0]);

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));
      xPropertySet.setPropertyValue("StringItemList", itemList);
      xPropertySet.setPropertyValue("SelectedItems", selectedItems);

      xPropertySet.setPropertyValue("MultiSelection", Boolean.FALSE);
      xPropertySet.setPropertyValue("Dropdown", Boolean.TRUE);

      dlgNameContainer.insertByName(componentID, ListBoxModel);

      dlgArabicListBox = (XListBox) UnoRuntime.queryInterface(XListBox.class,
          dlgControlContainer.getControl(componentID));

      dlgArabicListBox.addItemListener(new XItemListener() {
        @Override
        public void disposing(EventObject arg0) {}

        public void itemStateChanged(ItemEvent arg0) {
          selectedArbcLngg = getItemLanguague(dlgArabicListBox.getSelectedItem());
          selectedArbcVrsn = getItemVersion(dlgArabicListBox.getSelectedItem());
        }

      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertArabicNumberButton(String componentID, String label, int posx, int posy,
      int height, int width, boolean enabled, int tabIndex) {
    try {
      Object ButtonModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlButtonModel");

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, ButtonModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Label", label);
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("Toggle", Boolean.valueOf(true));
      xPropertySet.setPropertyValue("State", Short.valueOf((short) 1));
      xPropertySet.setPropertyValue("Align", Short.valueOf((short) 1));
      xPropertySet.setPropertyValue("FontRelief", com.sun.star.text.FontRelief.EMBOSSED);

      dlgNameContainer.insertByName(componentID, ButtonModel);
      dlgArabicNumberButton = (XButton) UnoRuntime.queryInterface(XButton.class,
          dlgControlContainer.getControl(componentID));


      dlgArabicNumberButton.addActionListener(new XActionListener() {

        @Override
        public void actionPerformed(ActionEvent arg0) {

        }

        @Override
        public void disposing(EventObject arg0) {}

      });

    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertArabicBoldButton(String componentID, String label, int posx, int posy,
      int height, int width, boolean enabled, int tabIndex) {
    try {
      Object ButtonModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlButtonModel");

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, ButtonModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Label", label);
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("Toggle", Boolean.valueOf(true));
      xPropertySet.setPropertyValue("State", Short.valueOf((short) 1));
      xPropertySet.setPropertyValue("Align", Short.valueOf((short) 1));
      xPropertySet.setPropertyValue("FontRelief", com.sun.star.text.FontRelief.EMBOSSED);

      dlgNameContainer.insertByName(componentID, ButtonModel);
      dlgArabicBoldButton = (XButton) UnoRuntime.queryInterface(XButton.class,
          dlgControlContainer.getControl(componentID));


      dlgArabicBoldButton.addActionListener(new XActionListener() {

        @Override
        public void actionPerformed(ActionEvent arg0) {

        }

        @Override
        public void disposing(EventObject arg0) {}

      });

    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertAyatAllCheckBox(String componentID, int posx, int posy, int height, int width,
      boolean enabled, int tabIndex) {
    try {
      Object checkBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, checkBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("State", Short.valueOf((short) 1));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("State", Short.valueOf((short) 1));

      selectedAllRngInd = shortToBoolean((short) 1);
      dlgNameContainer.insertByName(componentID, checkBoxModel);

      dlgAyatAllCheckBox =
          UnoRuntime.queryInterface(XCheckBox.class, dlgControlContainer.getControl(componentID));
      dlgAyatAllCheckBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          selectedAllRngInd = shortToBoolean(dlgAyatAllCheckBox.getState());

          if (!selectedAllRngInd) {
            dlgAyatToNumericField.setValue(selectedAytT);
            dlgAyatFromNumericField.setValue(selectedAytFrm);
          }

          enableComponent(AYATFROMLABELID, !selectedAllRngInd);
          enableComponent(AYATFROMNUMFLDID, !selectedAllRngInd);
          enableComponent(AYATTOLABELID, !selectedAllRngInd);
          enableComponent(AYATTONUMFLDID, !selectedAllRngInd);
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertAyatFromNumericField(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      Object NumericFieldModel =
          dlgMFactory.createInstance("com.sun.star.awt.UnoControlNumericFieldModel");

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, NumericFieldModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("Spin", Boolean.TRUE);
      xPropertySet.setPropertyValue("DecimalAccuracy", Short.valueOf((short) 0));

      dlgNameContainer.insertByName(componentID, NumericFieldModel);

      dlgAyatFromNumericField = UnoRuntime.queryInterface(XNumericField.class,
          dlgControlContainer.getControl(componentID));

      dlgAyatFromNumericField.setValue(selectedAytFrm);
      dlgAyatFromNumericField.setMin(selectedAytFrm);

      dlgAyaFromNumericSpinfield =
          UnoRuntime.queryInterface(XSpinField.class, dlgControlContainer.getControl(componentID));
      dlgAyaFromNumericSpinfield.addSpinListener(new XSpinListener() {
        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void down(SpinEvent arg0) {
          selectedAytFrm = Math.round(dlgAyatFromNumericField.getValue());
        }

        @Override
        public void first(SpinEvent arg0) {}

        @Override
        public void last(SpinEvent arg0) {}

        @Override
        public void up(SpinEvent arg0) {
          selectedAytFrm = Math.round(dlgAyatFromNumericField.getValue());
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertAyatToNumericField(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      Object NumericFieldModel =
          dlgMFactory.createInstance("com.sun.star.awt.UnoControlNumericFieldModel");


      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, NumericFieldModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("Spin", Boolean.TRUE);
      xPropertySet.setPropertyValue("DecimalAccuracy", Short.valueOf((short) 0));

      dlgNameContainer.insertByName(componentID, NumericFieldModel);

      dlgAyatToNumericField = UnoRuntime.queryInterface(XNumericField.class,
          dlgControlContainer.getControl(componentID));

      dlgAyatToNumericField.setValue(selectedAytT);
      dlgAyatToNumericField.setMax(selectedAytT);

      dlgAyaToNumericSpinfield =
          UnoRuntime.queryInterface(XSpinField.class, dlgControlContainer.getControl(componentID));
      dlgAyaToNumericSpinfield.addSpinListener(new XSpinListener() {
        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void down(SpinEvent arg0) {
          selectedAytT = Math.round(dlgAyatToNumericField.getValue());
        }

        @Override
        public void first(SpinEvent arg0) {}

        @Override
        public void last(SpinEvent arg0) {}

        @Override
        public void up(SpinEvent arg0) {
          selectedAytT = Math.round(dlgAyatToNumericField.getValue());
          System.out.printf("Selected Ayat To: %s\n", selectedAytT);
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertGroupBox(String componentID, String label, int posx, int posy, int height,
      int width, boolean enabled) {
    try {
      Object groupBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlGroupBoxModel");

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, groupBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("Label", label);
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      dlgNameContainer.insertByName(componentID, groupBoxModel);
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertLabel(String componentID, String label, short alignment, int posx, int posy,
      int height, int width, boolean enabled) {
    try {
      Object fixedTextModel =
          dlgMFactory.createInstance("com.sun.star.awt.UnoControlFixedTextModel");

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, fixedTextModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("Label", label);
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("Align", alignment);

      dlgNameContainer.insertByName(componentID, fixedTextModel);
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertLineByLineCheckBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      Object checkBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, checkBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("State", Short.valueOf((short) 1));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("State", Short.valueOf((short) 1));

      selectedLbLInd = shortToBoolean((short) 1);
      dlgNameContainer.insertByName(componentID, checkBoxModel);

      dlgLbLCheckBox =
          UnoRuntime.queryInterface(XCheckBox.class, dlgControlContainer.getControl(componentID));
      dlgLbLCheckBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          selectedLbLInd = shortToBoolean(dlgLbLCheckBox.getState());
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertNonArabicFontListBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      Object ListBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      selectedNnArbcFntNm = getDefaultNonArabicFontName();

      List<String> list = new ArrayList<String>();
      Locale locale = new Locale.Builder().setScript("LATN").build();

      short[] selectedItems = new short[1];

      String[] fonts =
          GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames(locale);
      for (int i = 0; i < fonts.length; i++) {
        if (new Font(fonts[i], Font.PLAIN, 10).canDisplay(A)) {
          list.add(fonts[i]);
          if (fonts[i].equals(selectedNnArbcFntNm)) {
            selectedItems[0] = (short) (list.size() - 1);
          }
        }
      }

      String[] itemList = list.toArray(new String[0]);

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));
      xPropertySet.setPropertyValue("StringItemList", itemList);
      xPropertySet.setPropertyValue("SelectedItems", selectedItems);

      xPropertySet.setPropertyValue("MultiSelection", Boolean.FALSE);
      xPropertySet.setPropertyValue("Dropdown", Boolean.TRUE);

      dlgNameContainer.insertByName(componentID, ListBoxModel);

      dlgNonArabicFontListBox = (XListBox) UnoRuntime.queryInterface(XListBox.class,
          dlgControlContainer.getControl(componentID));

      dlgNonArabicFontListBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {

        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertNonArabicFontSizeNumericField(String componentID, int posx, int posy,
      int height, int width, boolean enabled, int tabIndex) {
    try {
      Object NumericFieldModel =
          dlgMFactory.createInstance("com.sun.star.awt.UnoControlNumericFieldModel");

      selectedNnArbcFntSz = getDefaultNonArabicCharHeight();

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, NumericFieldModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("Spin", Boolean.TRUE);
      xPropertySet.setPropertyValue("DecimalAccuracy", Short.valueOf((short) 1));

      dlgNameContainer.insertByName(componentID, NumericFieldModel);

      dlgNonArabicFontsizeNumericField = UnoRuntime.queryInterface(XNumericField.class,
          dlgControlContainer.getControl(componentID));

      dlgNonArabicFontsizeNumericField.setValue(defaultNonArabicCharHeight);
      dlgNonArabicFontsizeNumericField.setMin(1);

      dlgNonArabicFontsizeNumericSpinfield =
          UnoRuntime.queryInterface(XSpinField.class, dlgControlContainer.getControl(componentID));
      dlgNonArabicFontsizeNumericSpinfield.addSpinListener(new XSpinListener() {
        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void down(SpinEvent arg0) {
          selectedNnArbcFntSz = dlgNonArabicFontsizeNumericField.getValue();
        }

        @Override
        public void first(SpinEvent arg0) {}

        @Override
        public void last(SpinEvent arg0) {}

        @Override
        public void up(SpinEvent arg0) {
          selectedNnArbcFntSz = dlgNonArabicFontsizeNumericField.getValue();
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertOkButton(String componentID, String label, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      Object ButtonModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlButtonModel");

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, ButtonModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Label", label);
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      dlgNameContainer.insertByName(componentID, ButtonModel);
      dlgOkButton = (XButton) UnoRuntime.queryInterface(XButton.class,
          dlgControlContainer.getControl(componentID));


      dlgOkButton.addActionListener(new XActionListener() {

        @Override
        public void actionPerformed(ActionEvent arg0) {
          writeSurah(selectedSurhNo);
          dlgDialog.endExecute();
        }

        @Override
        public void disposing(EventObject arg0) {}

      });

    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertSurahListBox(String componentID, int posx, int posy, int height, int width,
      boolean enabled, int tabIndex) {
    try {
      Object ListBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      String[] itemList = new String[114];
      for (int i = 0; i < 114; i++) {
        itemList[i] = QuranReader.getSurahName(i) + " (" + (i + 1) + ")";
      }
      short[] selectedItems = new short[] {(short) 0};

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("MultiSelection", Boolean.FALSE);
      xPropertySet.setPropertyValue("Dropdown", Boolean.TRUE);
      xPropertySet.setPropertyValue("StringItemList", itemList);
      xPropertySet.setPropertyValue("SelectedItems", selectedItems);

      selectedSurhNo = 1;
      selectedAytFrm = 1;
      selectedAytT = QuranReader.getSurahSize(selectedSurhNo - 1);

      dlgNameContainer.insertByName(componentID, ListBoxModel);

      dlgSurahListBox =
          UnoRuntime.queryInterface(XListBox.class, dlgControlContainer.getControl(componentID));
      dlgSurahListBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          selectedSurhNo = (short) (dlgSurahListBox.getSelectedItemPos() + 1);

          selectedAytFrm = 1;
          dlgAyatFromNumericField.setValue(selectedAytFrm);
          dlgAyatFromNumericField.setMin(selectedAytFrm);

          selectedAytT = QuranReader.getSurahSize(selectedSurhNo - 1);
          dlgAyatToNumericField.setValue(selectedAytT);
          dlgAyatToNumericField.setMax(selectedAytT);
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertTranslationCheckBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      Object checkBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, checkBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("State", Short.valueOf((short) 1));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("State", Short.valueOf((short) 0));
      selTrnsltnInd = shortToBoolean(Short.valueOf((short) 0));

      dlgNameContainer.insertByName(componentID, checkBoxModel);

      dlgTrnsltnCheckBox =
          UnoRuntime.queryInterface(XCheckBox.class, dlgControlContainer.getControl(componentID));
      dlgTrnsltnCheckBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {

          selTrnsltnInd = shortToBoolean(dlgTrnsltnCheckBox.getState());

          enableComponent(LANGTRNSLTNLSTBXID, selTrnsltnInd);
          enableComponent(FONTNONARABICGRPBXID, selectedTrnsltrtnInd || selTrnsltnInd);
          enableComponent(FONTNONARABICLABELID, selectedTrnsltrtnInd || selTrnsltnInd);
          enableComponent(FONTNONARABICLSTBXID, selectedTrnsltrtnInd || selTrnsltnInd);
          enableComponent(FONTNONARABICFONTSIZEID, selectedTrnsltrtnInd || selTrnsltnInd);
          enableComponent(FONTNONARABICFONTSIZENUMFLDID, selectedTrnsltrtnInd || selTrnsltnInd);
          enableComponent(OKBTTNID, selectedArbcInd || selTrnsltnInd || selectedTrnsltrtnInd);
          enableComponent(MISCGRPBXID, selectedArbcInd || selTrnsltnInd || selectedTrnsltrtnInd);
          enableComponent(MISCLBLCHKBXID, selectedArbcInd || selTrnsltnInd || selectedTrnsltrtnInd);
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertTranslationListBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      Object ListBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      List<String> list = new ArrayList<String>();

      QuranReader.getAllQuranVersions().entrySet().stream().forEach((entry) -> {
        String currentKey = entry.getKey();
        String[] currentValue = entry.getValue();
        for (int i = 0; i < currentValue.length; i++) {
          if (!currentKey.equals("Arabic")) {
            list.add(currentKey + " (" + currentValue[i] + ")");
          }
        }
      });

      String[] itemList = list.toArray(new String[0]);

      short[] selectedItems = new short[] {(short) 0};

      selectedTrnsltnLngg = getItemLanguague(itemList[0]);
      selectedTrnsltnVrsn = getItemVersion(itemList[0]);

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));
      xPropertySet.setPropertyValue("StringItemList", itemList);
      xPropertySet.setPropertyValue("SelectedItems", selectedItems);

      xPropertySet.setPropertyValue("MultiSelection", Boolean.FALSE);
      xPropertySet.setPropertyValue("Dropdown", Boolean.TRUE);

      dlgNameContainer.insertByName(componentID, ListBoxModel);

      dlgTrnsltnListBox =
          UnoRuntime.queryInterface(XListBox.class, dlgControlContainer.getControl(componentID));

      dlgTrnsltnListBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          selectedTrnsltnLngg = getItemLanguague(dlgTrnsltnListBox.getSelectedItem());
          selectedTrnsltnVrsn = getItemVersion(dlgTrnsltnListBox.getSelectedItem());
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertTransliterationCheckBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      Object checkBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlCheckBoxModel");

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, checkBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("State", Short.valueOf((short) 1));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));

      xPropertySet.setPropertyValue("State", Short.valueOf((short) 0));

      selectedAllRngInd = shortToBoolean((short) 1);
      dlgNameContainer.insertByName(componentID, checkBoxModel);

      dlgTrnsltrtnCheckBox =
          UnoRuntime.queryInterface(XCheckBox.class, dlgControlContainer.getControl(componentID));
      dlgTrnsltrtnCheckBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          selectedTrnsltrtnInd = shortToBoolean(dlgTrnsltrtnCheckBox.getState());

          enableComponent(LANGTRNSLTRTNLSTBXID, selectedTrnsltrtnInd);
          enableComponent(FONTNONARABICGRPBXID, selectedTrnsltrtnInd || selTrnsltnInd);
          enableComponent(FONTNONARABICLABELID, selectedTrnsltrtnInd || selTrnsltnInd);
          enableComponent(FONTNONARABICLSTBXID, selectedTrnsltrtnInd || selTrnsltnInd);
          enableComponent(FONTNONARABICFONTSIZEID, selectedTrnsltrtnInd || selTrnsltnInd);
          enableComponent(FONTNONARABICFONTSIZENUMFLDID, selectedTrnsltrtnInd || selTrnsltnInd);
          enableComponent(OKBTTNID, selectedArbcInd || selTrnsltnInd || selectedTrnsltrtnInd);
          enableComponent(MISCGRPBXID, selectedArbcInd || selTrnsltnInd || selectedTrnsltrtnInd);
          enableComponent(MISCLBLCHKBXID, selectedArbcInd || selTrnsltnInd || selectedTrnsltrtnInd);
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void insertTransliterationListBox(String componentID, int posx, int posy, int height,
      int width, boolean enabled, int tabIndex) {
    try {
      Object ListBoxModel = dlgMFactory.createInstance("com.sun.star.awt.UnoControlListBoxModel");

      List<String> list = new ArrayList<String>();

      QuranReader.getAllQuranVersions().entrySet().stream().forEach((entry) -> {
        String currentKey = entry.getKey();
        String[] currentValue = entry.getValue();
        for (int i = 0; i < currentValue.length; i++) {
          if (currentKey.equals("Transliteration")) {
            list.add(currentKey + " (" + currentValue[i] + ")");
          }
        }
      });

      String[] itemList = list.toArray(new String[0]);

      short[] selectedItems = new short[] {(short) 0};

      selectedTrnsltrtnLngg = getItemLanguague(itemList[0]);
      selectedTrnsltrtnVrsn = getItemVersion(itemList[0]);

      XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, ListBoxModel);
      xPropertySet.setPropertyValue("PositionX", Integer.valueOf(posx));
      xPropertySet.setPropertyValue("PositionY", Integer.valueOf(posy));
      xPropertySet.setPropertyValue("Width", Integer.valueOf(width));
      xPropertySet.setPropertyValue("Height", Integer.valueOf(height));
      xPropertySet.setPropertyValue("Name", componentID);
      xPropertySet.setPropertyValue("TabIndex", Short.valueOf((short) tabIndex));
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));
      xPropertySet.setPropertyValue("StringItemList", itemList);
      xPropertySet.setPropertyValue("SelectedItems", selectedItems);

      xPropertySet.setPropertyValue("MultiSelection", Boolean.FALSE);
      xPropertySet.setPropertyValue("Dropdown", Boolean.TRUE);

      dlgNameContainer.insertByName(componentID, ListBoxModel);

      dlgTrnsltrtnListBox =
          UnoRuntime.queryInterface(XListBox.class, dlgControlContainer.getControl(componentID));

      dlgTrnsltrtnListBox.addItemListener(new XItemListener() {

        @Override
        public void disposing(EventObject arg0) {}

        @Override
        public void itemStateChanged(ItemEvent arg0) {
          selectedTrnsltrtnLngg = getItemLanguague(dlgTrnsltrtnListBox.getSelectedItem());
          selectedTrnsltrtnVrsn = getItemVersion(dlgTrnsltrtnListBox.getSelectedItem());
        }
      });
    } catch (com.sun.star.uno.Exception e) {
      e.printStackTrace();
    }
  }

  public void writeSurah(int surahNumber) {
    XTextDocument textDoc = DocumentHelper.getCurrentDocument(dlgComponentContext);
    XController controller = textDoc.getCurrentController();
    XTextViewCursorSupplier textViewCursorSupplier = DocumentHelper.getCursorSupplier(controller);
    XTextViewCursor textViewCursor = textViewCursorSupplier.getViewCursor();
    XText text = textViewCursor.getText();
    XTextCursor textCursor = text.createTextCursorByRange(textViewCursor.getStart());
    XParagraphCursor paragraphCursor =
        UnoRuntime.queryInterface(XParagraphCursor.class, textCursor);
    XPropertySet paragraphCursorPropertySet = DocumentHelper.getPropertySet(paragraphCursor);

    try {
      paragraphCursorPropertySet.setPropertyValue("CharFontName", selectedNnArbcFntNm);
      paragraphCursorPropertySet.setPropertyValue("CharFontNameComplex", selectedArbcFntNm);
      paragraphCursorPropertySet.setPropertyValue("CharHeight", selectedNnArbcFntSz);
      paragraphCursorPropertySet.setPropertyValue("CharHeightComplex", selectedArbcFntSz);

      int from = (selectedAllRngInd) ? 1 : (int) selectedAytFrm;
      int to = (selectedAllRngInd) ? QuranReader.getSurahSize(surahNumber - 1) + 1
          : (int) selectedAytT + 1;

      if (selectedLbLInd) {
        String linea = " ", lineb = "", linec = "";
        for (int l = from; l < to; l++) {
          if (selectedArbcInd) {
            linea = linea + getAyahLine(surahNumber, l, selectedArbcLngg, selectedArbcVrsn) + "\n";
          }
          if (selTrnsltnInd) {
            lineb = lineb + getAyahLine(surahNumber, l, selectedTrnsltnLngg, selectedTrnsltnVrsn)
                + "\n";
          }
          if (selectedTrnsltrtnInd) {
            linec = linec
                + getAyahLine(surahNumber, l, selectedTrnsltrtnLngg, selectedTrnsltrtnVrsn) + "\n";
          }
        }
        addParagraph(text, paragraphCursor, linea + "\n", com.sun.star.style.ParagraphAdjust.RIGHT,
            com.sun.star.text.WritingMode2.RL_TB);
        addParagraph(text, paragraphCursor, lineb + "\n", com.sun.star.style.ParagraphAdjust.LEFT,
            com.sun.star.text.WritingMode2.LR_TB);
        addParagraph(text, paragraphCursor, linec + "\n", com.sun.star.style.ParagraphAdjust.LEFT,
            com.sun.star.text.WritingMode2.LR_TB);
      } else {
        if (selectedArbcInd) {
          String linea = "";
          for (int l = from; l < to; l++) {
            linea = linea + getAyahLine(surahNumber, l, selectedArbcLngg, selectedArbcVrsn) + " ";
          }
          addParagraph(text, paragraphCursor, linea + "\n",
              com.sun.star.style.ParagraphAdjust.RIGHT, com.sun.star.text.WritingMode2.RL_TB);
        }
        if (selTrnsltnInd) {
          String lineb = "";
          for (int l = from; l < to; l++) {
            lineb =
                lineb + getAyahLine(surahNumber, l, selectedTrnsltnLngg, selectedTrnsltnVrsn) + " ";
          }
          addParagraph(text, paragraphCursor, lineb + "\n", com.sun.star.style.ParagraphAdjust.LEFT,
              com.sun.star.text.WritingMode2.LR_TB);
        }
        if (selectedTrnsltrtnInd) {
          String linec = "";
          for (int l = from; l < to; l++) {
            linec = linec
                + getAyahLine(surahNumber, l, selectedTrnsltrtnLngg, selectedTrnsltrtnVrsn) + " ";
          }
          addParagraph(text, paragraphCursor, linec + "\n", com.sun.star.style.ParagraphAdjust.LEFT,
              com.sun.star.text.WritingMode2.LR_TB);
        }
      }
    } catch (com.sun.star.lang.IllegalArgumentException | UnknownPropertyException
        | PropertyVetoException | WrappedTargetException e) {
      e.printStackTrace();
    }
  }

}

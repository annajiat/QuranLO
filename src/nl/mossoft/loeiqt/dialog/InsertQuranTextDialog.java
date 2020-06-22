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

import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogEventHandler;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XNumericField;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.frame.XController;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.style.ParagraphAdjust;
import com.sun.star.text.ControlCharacter;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import java.awt.Font;
import java.awt.GraphicsEnvironment;
import java.util.List;
import java.util.Locale;
import nl.mossoft.loeiqt.helper.DialogHelper;
import nl.mossoft.loeiqt.helper.DocumentHelper;
import nl.mossoft.loeiqt.helper.QuranReader;

public class InsertQuranTextDialog implements XDialogEventHandler {

  private static final char ALIF = '\u0627';
  private static final String EVENTALLRANGEINDICATORCHANGED = "EvtAllRangeIndicatorChanged";
  private static final String EVENTARABICFONTNAMECHANGED = "EvtArabicFontNameChanged";
  private static final String EVENTARABICFONTSIZECHANGED = "EvtArabicFontSizeChanged";
  private static final String EVENTARABICNEWLINEINDICATORCHANGED =
      "EvtArabicNewLineIndicatorChanged";
  private static final String EVENTAYAHFROMCHANGED = "EvtAyahFromChanged";
  private static final String EVENTAYAHTOCHANGED = "EvtAyahToChanged";
  private static final String EVENTNONARABICFONTNAMECHANGED = "EvtNonArabicFontNameChanged";
  private static final String EVENTNONARABICFONTSIZECHANGED = "EvtNonArabicFontSizeChanged";
  private static final String EVENTNONARABICLANGUAGECHANGED = "EvtNonArabicLanguageChanged";
  private static final String EVENTNONARABICNEWLINEINDICATORCHANGED =
      "EvtNonArabicNewLineIndicatorChanged";
  private static final String EVENTOKBUTTONPRESSED = "EvtOkButtonPressed";
  private static final String EVENTSURAHCHANGED = "EvtSurahChanged";

  private static final String EVENTTRANSLATIONINDICATORCHANGED = "EvtTranslationIndicatorChanged";
  private static final String LPAR = "\uFD3E";

  private static final String RPAR = "\uFD3F";
  private static final char A = '\u0061';
  
  private static short booleanToShort(boolean b) {
    return (short) (b ? 1 : 0);
  }

  private static String numToArabNum(int n) {
    String[] arabNum = {"\u0660", "\u0661", "\u0662", "\u0663", "\u0664", "\u0665", "\u0666",
        "\u0667", "\u0668", "\u0669"};

    StringBuilder as = new StringBuilder();
    while (n > 0) {
      as.append(arabNum[n % 10]);
      n = n / 10;
    }
    return as.reverse().toString();
  }

  private static boolean shortToBoolean(short s) {
    return s != 0;
  }

  private float defaultArabicCharHeight;
  private String defaultArabicFontName;
  private float defaultNonArabicCharHeight;
  private String defaultNonArabicFontName;

  private XCheckBox dlgAllRangeCheckBox;
  private XListBox dlgArabicFontListBox;
  private XNumericField dlgArabicFontSizeNumericField;
  private XCheckBox dlgArabicNewlineCheckBox;
  private XNumericField dlgAyatFromNumericField;
  private XNumericField dlgAyatToNumericField;
  private XComponentContext dlgContext;
  private XDialog dlgDialog;
  private XListBox dlgNonArabicFontListBox;
  private XNumericField dlgNonArabicFontSizeNumericField;
  private XListBox dlgNonArabicLanguageListBox;
  private XCheckBox dlgNonArabicNewlineCheckBox;
  private XListBox dlgSurahListBox;
  private XCheckBox dlgTranslationCheckBox;

  private boolean selectedAllRangeIndicator;
  private String selectedArabicFontName;
  private double selectedArabicFontSize;
  private boolean selectedArabicNewlineIndicator;
  private int selectedAyatFrom;
  private int selectedAyatTo;
  private String selectedNonArabicFontName;
  private double selectedNonArabicFontSize;
  private String selectedNonArabicLanguage;
  private boolean selectedNonArabicNewlineIndicator;
  private String selectedNonArabicVersion;
  private short selectedSurahNum;
  private boolean selectedTranslationIndicator;

  private String[] supportedActions =
      new String[] {EVENTALLRANGEINDICATORCHANGED, EVENTAYAHFROMCHANGED, EVENTAYAHTOCHANGED,
          EVENTOKBUTTONPRESSED, EVENTARABICNEWLINEINDICATORCHANGED, EVENTSURAHCHANGED,
          EVENTARABICFONTSIZECHANGED, EVENTARABICFONTNAMECHANGED, EVENTNONARABICLANGUAGECHANGED,
          EVENTNONARABICFONTNAMECHANGED, EVENTNONARABICFONTSIZECHANGED,
          EVENTNONARABICNEWLINEINDICATORCHANGED, EVENTTRANSLATIONINDICATORCHANGED};

  public InsertQuranTextDialog(XComponentContext context) {
    dlgDialog = DialogHelper.createDialog("InsertQuranTextDialog.xdl", context, this);
    dlgContext = context;
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

  @Override
  public boolean callHandlerMethod(XDialog dialog, Object eventObject, String methodName)
      throws WrappedTargetException {

    if (methodName.equals(EVENTALLRANGEINDICATORCHANGED)) {
      dlgHandlerAllRangeCheckBox();
      return true;
    } else if (methodName.equals(EVENTAYAHFROMCHANGED)) {
      dlgHandlerAyahFromNumericField();
      return true;
    } else if (methodName.equals(EVENTAYAHTOCHANGED)) {
      dlgHandlerAyahToNumericField();
      return true;
    } else if (methodName.equals(EVENTOKBUTTONPRESSED)) {
      dlgHandlerOkButton();
      return true;
    } else if (methodName.equals(EVENTARABICNEWLINEINDICATORCHANGED)) {
      dlgHandlerArabicNewLineCheckbox();
      return true;
    } else if (methodName.equals(EVENTARABICFONTNAMECHANGED)) {
      dlgHandlerArabicFontListBox();
      return true;
    } else if (methodName.equals(EVENTARABICFONTSIZECHANGED)) {
      dlgHandlerArabicFontSizeNumericField();
      return true;
    } else if (methodName.equals(EVENTSURAHCHANGED)) {
      dlgHandlerSurahListBox();
      return true;
    } else if (methodName.equals(EVENTNONARABICLANGUAGECHANGED)) {
      dlgHandlerNonArabicLanguageListBox();
      return true;
    } else if (methodName.equals(EVENTNONARABICFONTNAMECHANGED)) {
      dlgHandlerNonArabicFontListBox();
      return true;
    } else if (methodName.equals(EVENTNONARABICFONTSIZECHANGED)) {
      dlgHandlerNonArabicFontSizeNumericField();
      return true;
    } else if (methodName.equals(EVENTNONARABICNEWLINEINDICATORCHANGED)) {
      dlgHandlerNonArabicNewLineCheckbox();
      return true;
    } else if (methodName.equals(EVENTTRANSLATIONINDICATORCHANGED)) {
      dlgHandlerTranslationCheckBox();
      return true;
    }

    return false;
  }

  private void configureDialogOnOpening() {
    getLoDocumentDefaults();
    initializeSurahListBox();
    initializeAyatGroupBox();
    initializeAllRangeCheckBox();
    initializeArabicFontListBox();
    initializeArabicFontSizeNumericField();
    initializeArabicNewlineCheckBox();
    initializeNonArabicLanguageListBox();
    initializeNonArabicFontListBox();
    initializeNonArabicFontSizeNumericField();
    initializeNonArabicNewlineCheckBox();
    initializeTranslationCheckBox();
  }

  private void createOutput() {
    XTextDocument textDoc = DocumentHelper.getCurrentDocument(dlgContext);
    XController controller = textDoc.getCurrentController();
    XTextViewCursorSupplier textViewCursorSupplier = DocumentHelper.getCursorSupplier(controller);
    XTextViewCursor textViewCursor = textViewCursorSupplier.getViewCursor();
    XText text = textViewCursor.getText();
    XTextCursor textCursor = text.createTextCursorByRange(textViewCursor.getStart());
    XParagraphCursor paragraphCursor =
        UnoRuntime.queryInterface(XParagraphCursor.class, textCursor);
    XPropertySet paragraphCursorPropertySet = DocumentHelper.getPropertySet(paragraphCursor);

    try {
      paragraphCursorPropertySet.setPropertyValue("CharFontName", selectedNonArabicFontName);
      paragraphCursorPropertySet.setPropertyValue("CharFontNameComplex", selectedArabicFontName);
      paragraphCursorPropertySet.setPropertyValue("CharHeight", selectedNonArabicFontSize);
      paragraphCursorPropertySet.setPropertyValue("CharHeightComplex", selectedArabicFontSize);

      if (selectedTranslationIndicator) {
        QuranReader ar = new QuranReader("Arabic", "Medina", dlgContext);
        QuranReader tr =
            new QuranReader(selectedNonArabicLanguage, selectedNonArabicVersion, dlgContext);

        if (selectedAllRangeIndicator) {
          List<String> ayatAr = ar.getAllAyatOfSuraNo(selectedSurahNum);
          List<String> ayatTr = tr.getAllAyatOfSuraNo(selectedSurahNum);

          if (selectedArabicNewlineIndicator) {
            if (selectedSurahNum != 1 && selectedSurahNum != 9) {
              addParagraph(text, paragraphCursor, ar.getBismillah(),
                  com.sun.star.style.ParagraphAdjust.RIGHT, com.sun.star.text.WritingMode2.RL_TB);
              addParagraph(text, paragraphCursor, tr.getBismillah(),
                  com.sun.star.style.ParagraphAdjust.LEFT, com.sun.star.text.WritingMode2.LR_TB);
            }

            for (int i = 0; i < ayatAr.size(); i++) {
              addParagraph(text, paragraphCursor,
                  ayatAr.get(i) + " " + RPAR + numToArabNum(i + 1) + LPAR + " ",
                  com.sun.star.style.ParagraphAdjust.RIGHT, com.sun.star.text.WritingMode2.RL_TB);
              addParagraph(text, paragraphCursor, "(" + (i + 1) + ") " + ayatTr.get(i),
                  com.sun.star.style.ParagraphAdjust.LEFT, com.sun.star.text.WritingMode2.LR_TB);
            }
          } else {
            String tlnl = (selectedNonArabicNewlineIndicator) ? "\n" : "";
            String lineAr =
                (selectedSurahNum != 1 && selectedSurahNum != 9) ? ar.getBismillah() + " " : " ";
            String lineTr =
                (selectedSurahNum != 1 && selectedSurahNum != 9) ? tr.getBismillah() + " " : tlnl;
            for (int i = 0; i < ayatAr.size(); i++) {
              lineAr = lineAr + ayatAr.get(i) + " " + RPAR + numToArabNum(i + 1) + LPAR + " ";
              lineTr = lineTr + "(" + (i + 1) + ") " + ayatTr.get(i) + " " + tlnl;
            }
            addParagraph(text, paragraphCursor, lineAr, com.sun.star.style.ParagraphAdjust.RIGHT,
                com.sun.star.text.WritingMode2.RL_TB);
            addParagraph(text, paragraphCursor, lineTr, com.sun.star.style.ParagraphAdjust.LEFT,
                com.sun.star.text.WritingMode2.LR_TB);
          }
        } else {
          List<String> ayatAr =
              ar.getAyatFromToOfSuraNo(selectedSurahNum, selectedAyatFrom, selectedAyatTo);
          List<String> ayatTr =
              tr.getAyatFromToOfSuraNo(selectedSurahNum, selectedAyatFrom, selectedAyatTo);

          if (selectedArabicNewlineIndicator) {
            for (int i = 0; i < ayatAr.size(); i++) {
              addParagraph(text, paragraphCursor,
                  ayatAr.get(i) + " " + RPAR + numToArabNum(selectedAyatFrom + i) + LPAR + " ",
                  com.sun.star.style.ParagraphAdjust.RIGHT, com.sun.star.text.WritingMode2.RL_TB);
              addParagraph(text, paragraphCursor,
                  "(" + (selectedAyatFrom + i) + ") " + ayatTr.get(i) + " ",
                  com.sun.star.style.ParagraphAdjust.LEFT, com.sun.star.text.WritingMode2.LR_TB);
            }
          } else {
            String tln1 = (selectedNonArabicNewlineIndicator) ? "\n" : "";
            String lineAr = "";
            String lineTr = "";
            for (int i = 0; i < ayatAr.size(); i++) {
              lineAr = lineAr + ayatAr.get(i) + " " + RPAR + numToArabNum(selectedAyatFrom + i)
                  + LPAR + " ";
              lineTr = lineTr + "(" + (selectedAyatFrom + i) + ") " + ayatTr.get(i) + " " + tln1;
            }
            addParagraph(text, paragraphCursor, lineAr, com.sun.star.style.ParagraphAdjust.RIGHT,
                com.sun.star.text.WritingMode2.RL_TB);
            addParagraph(text, paragraphCursor, lineTr, com.sun.star.style.ParagraphAdjust.LEFT,
                com.sun.star.text.WritingMode2.LR_TB);
          }
        }

      } else {
        QuranReader ar = new QuranReader("Arabic", "Medina", dlgContext);

        if (selectedAllRangeIndicator) {
          List<String> ayatAr = ar.getAllAyatOfSuraNo(selectedSurahNum);

          if (selectedArabicNewlineIndicator) {
            if (selectedSurahNum != 1 && selectedSurahNum != 9) {
              addParagraph(text, paragraphCursor, ar.getBismillah(),
                  com.sun.star.style.ParagraphAdjust.RIGHT, com.sun.star.text.WritingMode2.RL_TB);
            }

            for (int i = 0; i < ayatAr.size(); i++) {
              addParagraph(text, paragraphCursor,
                  ayatAr.get(i) + " " + RPAR + numToArabNum(i + 1) + LPAR + " ",
                  com.sun.star.style.ParagraphAdjust.RIGHT, com.sun.star.text.WritingMode2.RL_TB);
            }
          } else {
            String lineAr =
                (selectedSurahNum != 1 && selectedSurahNum != 9) ? ar.getBismillah() + " " : " ";
            for (int i = 0; i < ayatAr.size(); i++) {
              lineAr = lineAr + ayatAr.get(i) + " " + RPAR + numToArabNum(i + 1) + LPAR + " ";
            }
            addParagraph(text, paragraphCursor, lineAr, com.sun.star.style.ParagraphAdjust.RIGHT,
                com.sun.star.text.WritingMode2.RL_TB);
          }
        } else {
          List<String> ayatAr =
              ar.getAyatFromToOfSuraNo(selectedSurahNum, selectedAyatFrom, selectedAyatTo);

          if (selectedArabicNewlineIndicator) {
            for (int i = 0; i < ayatAr.size(); i++) {
              addParagraph(text, paragraphCursor,
                  ayatAr.get(i) + " " + RPAR + numToArabNum(selectedAyatFrom + i) + LPAR + " ",
                  com.sun.star.style.ParagraphAdjust.RIGHT, com.sun.star.text.WritingMode2.RL_TB);
            }
          } else {
            String lineAr =
                (selectedSurahNum != 1 && selectedSurahNum != 9) ? ar.getBismillah() + " " : "";
            for (int i = 0; i < ayatAr.size(); i++) {
              lineAr = lineAr + ayatAr.get(i) + " " + RPAR + numToArabNum(selectedAyatFrom + i)
                  + LPAR + " ";
            }
            addParagraph(text, paragraphCursor, lineAr, com.sun.star.style.ParagraphAdjust.RIGHT,
                com.sun.star.text.WritingMode2.RL_TB);
          }
        }
      }
    } catch (com.sun.star.lang.IllegalArgumentException | UnknownPropertyException
        | PropertyVetoException | WrappedTargetException e) {
      e.printStackTrace();
    }
  }

  private void dlgHandlerAllRangeCheckBox() {
    selectedAllRangeIndicator = shortToBoolean(dlgAllRangeCheckBox.getState());

    enableAyatGroupBox(!selectedAllRangeIndicator);
  }

  private void dlgHandlerArabicFontListBox() {
    selectedArabicFontName = dlgArabicFontListBox.getSelectedItem();
  }

  private void dlgHandlerArabicFontSizeNumericField() {
    selectedArabicFontSize = dlgArabicFontSizeNumericField.getValue();
  }

  private void dlgHandlerArabicNewLineCheckbox() {
    selectedArabicNewlineIndicator = shortToBoolean(dlgArabicNewlineCheckBox.getState());
  }

  private void dlgHandlerAyahFromNumericField() {
    selectedAyatFrom = (int) dlgAyatFromNumericField.getValue();
  }

  private void dlgHandlerAyahToNumericField() {
    selectedAyatTo = (int) dlgAyatToNumericField.getValue();

  }

  private void dlgHandlerNonArabicFontListBox() {
    selectedNonArabicFontName = dlgNonArabicFontListBox.getSelectedItem();
  }

  private void dlgHandlerNonArabicFontSizeNumericField() {
    selectedNonArabicFontSize = dlgNonArabicFontSizeNumericField.getValue();
  }

  private void dlgHandlerNonArabicLanguageListBox() {
    String selectedNonArabicitem = dlgNonArabicLanguageListBox.getSelectedItem();

    String[] itemsSelected = selectedNonArabicitem.split("[(]");
    selectedNonArabicLanguage = itemsSelected[0].trim();
    selectedNonArabicVersion = itemsSelected[1].replace(")", " ").trim().replace(" ", "_");
  }

  private void dlgHandlerNonArabicNewLineCheckbox() {
    selectedNonArabicNewlineIndicator = shortToBoolean(dlgNonArabicNewlineCheckBox.getState());
  }

  private void dlgHandlerOkButton() {
    createOutput();
    dlgDialog.endExecute();
  }

  private void dlgHandlerSurahListBox() {
    selectedSurahNum = (short) (dlgSurahListBox.getSelectedItemPos() + 1);
  }

  private void dlgHandlerTranslationCheckBox() {
    selectedTranslationIndicator = shortToBoolean(dlgTranslationCheckBox.getState());

    enableNonArabicGroupBox(selectedTranslationIndicator);
  }

  private void enableAyatGroupBox(boolean enable) {
    try {
      DialogHelper.enableComponent(dlgDialog, "AyatFromLabel", enable);
      DialogHelper.enableComponent(dlgDialog, "AyatFromNumericField", enable);
      DialogHelper.enableComponent(dlgDialog, "AyatToLabel", enable);
      DialogHelper.enableComponent(dlgDialog, "AyatToNumericField", enable);

      if (enable) {
        selectedAyatFrom = 1;
        selectedAyatTo = QuranReader.getSurahSize(selectedSurahNum - 1);

        dlgAyatFromNumericField.setValue(selectedAyatFrom);
        dlgAyatToNumericField.setMax(selectedAyatTo);
        dlgAyatToNumericField.setValue(selectedAyatTo);
      } else {
        dlgAyatFromNumericField.setValue(1);
        dlgAyatToNumericField.setValue(286);
      }
    } catch (com.sun.star.lang.IllegalArgumentException | UnknownPropertyException
        | PropertyVetoException | WrappedTargetException e) {
      e.printStackTrace();
    }
  }

  private void enableNonArabicGroupBox(boolean enable) {
    try {
      DialogHelper.enableComponent(dlgDialog, "NonArabicGroupBox", enable);
      DialogHelper.enableComponent(dlgDialog, "NonArabicLanguageListBox", enable);
      DialogHelper.enableComponent(dlgDialog, "NonArabicLanguageLabel", enable);
      DialogHelper.enableComponent(dlgDialog, "NonArabicFontLabel", enable);
      DialogHelper.enableComponent(dlgDialog, "NonArabicFontListBox", enable);
      DialogHelper.enableComponent(dlgDialog, "NonArabicFontSizeLabel", enable);
      DialogHelper.enableComponent(dlgDialog, "NonArabicFontSizeNumericField", enable);
      DialogHelper.enableComponent(dlgDialog, "NonArabicNewLineLabel", enable);
      DialogHelper.enableComponent(dlgDialog, "NonArabicNewlineCheckBox", enable);
      if (enable) {
        dlgNonArabicNewlineCheckBox.setState(booleanToShort(selectedArabicNewlineIndicator));
      } else {
        dlgNonArabicNewlineCheckBox.setState(booleanToShort(false));
      }
    } catch (UnknownPropertyException | PropertyVetoException | WrappedTargetException e) {
      e.printStackTrace();
    }

  }

  private float getDefaultArabicCharHeight() {
    return defaultArabicCharHeight;
  }

  private String getDefaultArabicFontName() {
    return defaultArabicFontName;
  }

  private float getDefaultNonArabicCharHeight() {
    return defaultNonArabicCharHeight;
  }

  private String getDefaultNonArabicFontName() {
    return defaultNonArabicFontName;
  }

  private void getLoDocumentDefaults() {
    XTextDocument textDoc = DocumentHelper.getCurrentDocument(dlgContext);
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

  @Override
  public String[] getSupportedMethodNames() {
    return supportedActions;
  }

  private void initializeAllRangeCheckBox() {
    dlgAllRangeCheckBox = DialogHelper.getCheckbox(dlgDialog, "AllRangeCheckBox");

    dlgAllRangeCheckBox.setState(booleanToShort(true));
    selectedAllRangeIndicator = shortToBoolean(dlgAllRangeCheckBox.getState());

    enableAyatGroupBox(!selectedAllRangeIndicator);
  }

  private void initializeArabicFontListBox() {
    dlgArabicFontListBox = DialogHelper.getListBox(dlgDialog, "ArabicFontListBox");
    Locale locale = new Locale.Builder().setScript("ARAB").build();
    String[] fonts =
        GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames(locale);
    for (int i = 0; i < fonts.length; i++) {
      if (new Font(fonts[i], Font.PLAIN, 10).canDisplay(ALIF)) {
        dlgArabicFontListBox.addItem(fonts[i], (short) i);
      }
    }
    selectedArabicFontName = getDefaultArabicFontName();

    dlgArabicFontListBox.selectItem(selectedArabicFontName, true);
  }

  private void initializeArabicFontSizeNumericField() {
    dlgArabicFontSizeNumericField =
        DialogHelper.getNumericField(dlgDialog, "ArabicFontSizeNumericField");

    selectedArabicFontSize = getDefaultArabicCharHeight();

    dlgArabicFontSizeNumericField.setValue(selectedArabicFontSize);
  }

  private void initializeArabicNewlineCheckBox() {
    dlgArabicNewlineCheckBox = DialogHelper.getCheckbox(dlgDialog, "ArabicNewlineCheckBox");

    dlgArabicNewlineCheckBox.setState(booleanToShort(true));
    selectedArabicNewlineIndicator = shortToBoolean(dlgArabicNewlineCheckBox.getState());
  }

  private void initializeAyatGroupBox() {
    dlgAyatFromNumericField = DialogHelper.getNumericField(dlgDialog, "AyatFromNumericField");
    dlgAyatToNumericField = DialogHelper.getNumericField(dlgDialog, "AyatToNumericField");
  }

  private void initializeNonArabicFontListBox() {
    

    dlgNonArabicFontListBox = DialogHelper.getListBox(dlgDialog, "NonArabicFontListBox");

    Locale locale = new Locale.Builder().setScript("LATN").build();
    String[] fonts =
        GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames(locale);
    for (int i = 0; i < fonts.length; i++) {
      if (new Font(fonts[i], Font.PLAIN, 10).canDisplay(A)) {
        dlgNonArabicFontListBox.addItem(fonts[i], (short) i);
      }
    }
    selectedNonArabicFontName = getDefaultNonArabicFontName();

    dlgNonArabicFontListBox.selectItem(selectedNonArabicFontName, true);
  }

  private void initializeNonArabicFontSizeNumericField() {
    dlgNonArabicFontSizeNumericField =
        DialogHelper.getNumericField(dlgDialog, "NonArabicFontSizeNumericField");

    selectedNonArabicFontSize = getDefaultNonArabicCharHeight();

    dlgNonArabicFontSizeNumericField.setValue(selectedNonArabicFontSize);

  }

  private void initializeNonArabicLanguageListBox() {
    dlgNonArabicLanguageListBox = DialogHelper.getListBox(dlgDialog, "NonArabicLanguageListBox");

    QuranReader.getAllQuranVersions().entrySet().stream().forEach((entry) -> {
      String currentKey = entry.getKey();
      String[] currentValue = entry.getValue();
      for (int i = 0; i < currentValue.length; i++) {
        if (!currentKey.equals("Arabic")) {
          dlgNonArabicLanguageListBox.addItem(currentKey + " (" + currentValue[i] + ")", (short) i);
        }
      }
    });

    dlgNonArabicLanguageListBox.selectItemPos((short) 0, true);

    String selectedNonArabicitem = dlgNonArabicLanguageListBox.getSelectedItem();

    String[] itemsSelected = selectedNonArabicitem.split("[(]");
    selectedNonArabicLanguage = itemsSelected[0].trim();
    selectedNonArabicVersion = itemsSelected[1].replace(")", " ").trim().replace(" ", "_");
  }

  private void initializeNonArabicNewlineCheckBox() {
    dlgNonArabicNewlineCheckBox = DialogHelper.getCheckbox(dlgDialog, "NonArabicNewlineCheckBox");

    dlgNonArabicNewlineCheckBox.setState(booleanToShort(true));
    selectedNonArabicNewlineIndicator = shortToBoolean(dlgNonArabicNewlineCheckBox.getState());
  }

  private void initializeSurahListBox() {
    dlgSurahListBox = DialogHelper.getListBox(dlgDialog, "SurahListBox");

    for (int i = 0; i < 114; i++) {
      dlgSurahListBox.addItem(QuranReader.getSurahName(i) + " (" + (i + 1) + ")", (short) i);
    }
    dlgSurahListBox.selectItemPos((short) 0, true);
    selectedSurahNum = (short) (dlgSurahListBox.getSelectedItemPos() + 1);
  }

  private void initializeTranslationCheckBox() {
    dlgTranslationCheckBox = DialogHelper.getCheckbox(dlgDialog, "TranslationCheckBox");

    dlgTranslationCheckBox.setState(booleanToShort(false));
    selectedTranslationIndicator = shortToBoolean(dlgTranslationCheckBox.getState());

    enableNonArabicGroupBox(selectedTranslationIndicator);
  }

  public void show() {
    configureDialogOnOpening();
    dlgDialog.execute();
  }

}

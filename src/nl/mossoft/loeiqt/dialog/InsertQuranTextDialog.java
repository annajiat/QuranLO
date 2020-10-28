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
import com.sun.star.awt.XProgressBar;
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
import java.util.Locale;
import nl.mossoft.loeiqt.helper.DialogHelper;
import nl.mossoft.loeiqt.helper.DocumentHelper;
import nl.mossoft.loeiqt.helper.QuranReader;

public class InsertQuranTextDialog implements XDialogEventHandler {

  private static final String ARABIC = "Arabic";
  private static final String DIALOG_ALL_AYAT_CHECKBOX_ID = "AllAyatCheckBoxID";
  private static final String DIALOG_ARABIC_CHECKBOX_ID = "ArabicCheckBoxID";
  private static final String DIALOG_ARABIC_FONT_GROUPBOX_ID = "ArabicFontGroupBoxID";
  private static final String DIALOG_ARABIC_FONT_LABEL_ID = "ArabicFontLabelID";
  private static final String DIALOG_ARABIC_FONT_LISTBOX_ID = "ArabicFontListBoxID";
  private static final String DIALOG_ARABIC_FONTSIZE_LABEL_ID = "ArabicFontSizeLabelID";
  private static final String DIALOG_ARABIC_FONTSIZE_NUMERICFIELD_ID =
      "ArabicFontSizeNumericFieldID";
  private static final String DIALOG_ARABIC_LISTBOX_ID = "ArabicListBoxID";
  private static final String DIALOG_AYAT_FROM_NUMERIC_FIELD_ID = "AyatFromNumericFieldID";
  private static final String DIALOG_AYAT_TO_NUMERIC_FIELD_ID = "AyatToNumericFieldID";
  private static final String DIALOG_LINE_BY_LINE_CHECKBOX_ID = "LineByLineCheckBoxID";
  private static final String DIALOG_LINE_BY_LINE_LABEL_ID = "LineBylineLabelID";
  private static final String DIALOG_MISCELLANEOUS_GROUPBOX_ID = "MiscellaneousGroupBoxID";
  private static final String DIALOG_NON_ARABIC_FONT_GROUPBOX_ID = "NonArabicFontGroupBoxID";
  private static final String DIALOG_NON_ARABIC_FONT_LABEL_ID = "NonArabicFontLabelID";
  private static final String DIALOG_NON_ARABIC_FONT_LISTBOX_ID = "NonArabicFontListBoxID";
  private static final String DIALOG_NON_ARABIC_FONTSIZE_LABEL_ID = "NonArabicFontSizeLabelID";
  private static final String DIALOG_NON_ARABIC_FONTSIZE_NUMERICFIELD_ID =
      "NonArabicFontSizeNumericFieldID";
  private static final String DIALOG_OK_BUTTON_ID = "OkButtonID";
  private static final String DIALOG_SURAH_LISTBOX_ID = "SurahListBoxID";
  private static final String DIALOG_TRANSLATION_CHECKBOX_ID = "TranslationCheckBoxID";
  private static final String DIALOG_TRANSLATION_LISTBOX_ID = "TranslationListBoxID";
  private static final String DIALOG_TRANSLITERATION_CHECKBOX_ID = "TransliterationCheckBoxID";
  private static final String DIALOG_TRANSLITERATION_LISTBOX_ID = "TransliterationListBoxID";
  private static final String DIALOG_WRITE_SURAH_PROGRESSBAR = "WriteSurahProgressBar";
  private static final String LEFT_PARENTHESIS = new String(Character.toChars(0xFD3E));
  private static final String ON_ALL_AYAT_CHECKBUTTON_PRESSED = "onAllAyatCheckButtonPressed";
  private static final String ON_ARABIC_CHECKBUTTON_PRESSED = "onArabicCheckButtonPressed";
  private static final String ON_ARABIC_FONT_SELECTED = "onArabicFontSelected";
  private static final String ON_ARABIC_FONT_SIZE_TEXT_CHANGED = "onArabicFontSizeTextChanged";
  private static final String ON_ARABIC_VERSION_SELECTED = "onArabicVersionSelected";
  private static final String ON_AYAT_FROM_TEXT_CHANGED = "onAyatFromTextChanged";
  private static final String ON_AYAT_TO_TEXT_CHANGED = "onAyatToTextChanged";
  private static final String ON_LINE_BY_LINE_CHECKBUTTON_PRESSED =
      "onLineByLineCheckButtonPressed";
  private static final String ON_NON_ARABIC_FONT_SELECTED = "onNonArabicFontSelected";
  private static final String ON_NON_ARABIC_FONT_SIZE_TEXT_CHANGED =
      "onNonArabicFontSizeTextChanged";
  private static final String ON_OK_BUTTON_PRESSED = "onOkButtonPressed";
  private static final String ON_SURAH_SELECTED = "onSurahSelected";
  private static final String ON_TRANSLATION_CHECKBUTTON_PRESSED =
      "onTranslationCheckButtonPressed";
  private static final String ON_TRANSLATION_VERSION_SELECTED = "onTranslationVersionSelected";
  private static final String ON_TRANSLITERATION_CHECKBUTTON_PRESSED =
      "onTransliterationCheckButtonPressed";
  private static final String ON_TRANSLITERATION_VERSION_SELECTED =
      "onTransliterationVersionSelected";
  private static final String RIGHT_PARENTHESIS = new String(Character.toChars(0xFD3F));

  /**
   * Convert boolean to short.
   *
   * @param b the boolean
   * @return the short
   */
  private static short boolean2Short(final boolean b) {
    return (short) (b ? 1 : 0);
  }

  private static String getItemLanguague(final String item) {
    final String[] itemsSelected = item.split("[(]");
    return itemsSelected[0].trim();
  }

  private static String getItemVersion(final String item) {
    final String[] itemsSelected = item.split("[(]");
    return itemsSelected[1].replace(")", " ").trim().replace(" ", "_");
  }

  /**
   * Convert short to boolean.
   *
   * @param s the short
   * @return the boolean
   */
  private static boolean short2Boolean(final short s) {
    return s != 0;
  }

  private double defaultArabicCharHeight;
  private String defaultArabicFontName;
  private double defaultNonArabicCharHeight;
  private String defaultNonArabicFontName;
  private final XComponentContext dlgContext;
  private final XDialog dlgDialog;
  private boolean selectedAllAyatInd = true;
  private String selectedArabicFontName = "";
  private double selectedArabicFontSize;
  private boolean selectedArabicInd = false;
  private String selectedArabicLanguage = "";
  private String selectedArabicVersion = "";
  private int selectedAyatFrom = 1;
  private int selectedAyatTo = 7;
  private boolean selectedLineByLineInd = true;
  private boolean selectedLineNumberInd = true;
  private String selectedNonArabicFontName = "";
  private double selectedNonArabicFontSize;
  private int selectedSurahNo = 1;
  private boolean selectedTranslationInd = false;
  private String selectedTranslationLanguage = "";
  private String selectedTranslationVersion = "";
  private boolean selectedTransliterationInd = false;
  private String selectedTransliterationLanguage = "";
  private String selectedTransliterationVersion = "";
  private final String[] supportedActions =
      new String[] {ON_ALL_AYAT_CHECKBUTTON_PRESSED, ON_ARABIC_CHECKBUTTON_PRESSED,
          ON_ARABIC_VERSION_SELECTED, ON_ARABIC_FONT_SELECTED, ON_ARABIC_FONT_SIZE_TEXT_CHANGED,
          ON_AYAT_FROM_TEXT_CHANGED, ON_AYAT_TO_TEXT_CHANGED, ON_LINE_BY_LINE_CHECKBUTTON_PRESSED,
          ON_NON_ARABIC_FONT_SELECTED, ON_NON_ARABIC_FONT_SIZE_TEXT_CHANGED, ON_OK_BUTTON_PRESSED,
          ON_SURAH_SELECTED, ON_TRANSLATION_CHECKBUTTON_PRESSED, ON_TRANSLATION_VERSION_SELECTED,
          ON_TRANSLITERATION_CHECKBUTTON_PRESSED, ON_TRANSLITERATION_VERSION_SELECTED};

  public InsertQuranTextDialog(final XComponentContext context) {
    dlgDialog = DialogHelper.createDialog("InsertQuranTextDialog.xdl", context, this);
    dlgContext = context;
  }

  @Override
  public boolean callHandlerMethod(final XDialog dialog, final Object eventObject,
      final String methodName) throws WrappedTargetException {

    if (methodName.equals(ON_ALL_AYAT_CHECKBUTTON_PRESSED)) {
      handleAllAyatCheckButtonPressed();
      return true; // Event was handled
    } else if (methodName.equals(ON_ARABIC_CHECKBUTTON_PRESSED)) {
      handleArabicCheckButtonPressed();
      return true; // Event was handled
    } else if (methodName.equals(ON_ARABIC_VERSION_SELECTED)) {
      handleArabicVersionSelected();
      return true; // Event was handled
    } else if (methodName.equals(ON_ARABIC_FONT_SELECTED)) {
      handleArabicFontSelected();
      return true; // Event was handled
    } else if (methodName.equals(ON_ARABIC_FONT_SIZE_TEXT_CHANGED)) {
      handleArabicFontTextChanged();
      return true; // Event was handled
    } else if (methodName.equals(ON_AYAT_FROM_TEXT_CHANGED)) {
      handleAyatFromTextChanged();
      return true; // Event was handled
    } else if (methodName.equals(ON_AYAT_TO_TEXT_CHANGED)) {
      handleAyatToTextChanged();
      return true; // Event was handled
    } else if (methodName.equals(ON_LINE_BY_LINE_CHECKBUTTON_PRESSED)) {
      handleLineByLineCheckButtonPressed();
      return true; // Event was handled
    } else if (methodName.equals(ON_NON_ARABIC_FONT_SELECTED)) {
      handleNonArabicFontSelected();
      return true; // Event was handled
    } else if (methodName.equals(ON_NON_ARABIC_FONT_SIZE_TEXT_CHANGED)) {
      handleNonArabicFontTextChanged();
      return true; // Event was handled
    } else if (methodName.equals(ON_OK_BUTTON_PRESSED)) {
      handleOkButtonPressed();
      return true; // Event was handled
    } else if (methodName.equals(ON_SURAH_SELECTED)) {
      handleSurahSelected();
      return true; // Event was handled
    } else if (methodName.equals(ON_TRANSLATION_CHECKBUTTON_PRESSED)) {
      handleTranslationCheckButtonPressed();
      return true; // Event was handled
    } else if (methodName.equals(ON_TRANSLATION_VERSION_SELECTED)) {
      handleTranslationVersionSelected();
      return true; // Event was handled
    } else if (methodName.equals(ON_TRANSLITERATION_CHECKBUTTON_PRESSED)) {
      handleTransliterationCheckButtonPressedSelected();
      return true; // Event was handled
    } else if (methodName.equals(ON_TRANSLITERATION_VERSION_SELECTED)) {
      handleTransliterationVersionSelected();
      return true; // Event was handled
    }
    return false; // Event was not handled
  }

  private String getAyahLine(final int surahno, final int ayahno, final String language,
      final String version) {

    final QuranReader qr = new QuranReader(language, version, dlgContext);
    String line = ((ayahno == 1) && (surahno != 1 && surahno != 9))
        ? getBismillah(language, version) + " " + qr.getAyahNoOfSuraNo(surahno, ayahno)
        : qr.getAyahNoOfSuraNo(surahno, ayahno);

    if (selectedLineNumberInd) {
      if (language.equals(ARABIC)) {
        line = line + " " + RIGHT_PARENTHESIS
            + QuranReader.numToArabNum(ayahno, selectedArabicFontName) + LEFT_PARENTHESIS + " ";
      } else {
        line = "(" + ayahno + ") " + line;
      }
    }
    return line;
  }

  private String getBismillah(final String language, final String version) {
    final QuranReader qr = new QuranReader(language, version, dlgContext);
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
    final XTextDocument textDoc = DocumentHelper.getCurrentDocument(dlgContext);
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

  /**
   * Handler for the All Ayat CheckButtun.
   */
  private void handleAllAyatCheckButtonPressed() {
    final XCheckBox dlgAllAyatCheckbox =
        DialogHelper.getCheckBox(dlgDialog, DIALOG_ALL_AYAT_CHECKBOX_ID);
    dlgAllAyatCheckbox.setState(boolean2Short(!selectedAllAyatInd));
    selectedAllAyatInd = short2Boolean(dlgAllAyatCheckbox.getState());

    DialogHelper.enableButton(dlgDialog, DIALOG_AYAT_FROM_NUMERIC_FIELD_ID, !selectedAllAyatInd);
    DialogHelper.enableButton(dlgDialog, DIALOG_AYAT_TO_NUMERIC_FIELD_ID, !selectedAllAyatInd);

    if (selectedAllAyatInd) {
      final XNumericField dlgAyatFromNumericField =
          DialogHelper.getNumericField(dlgDialog, DIALOG_AYAT_FROM_NUMERIC_FIELD_ID);
      dlgAyatFromNumericField.setValue(1);
      selectedAyatFrom = (int) Math.round(dlgAyatFromNumericField.getValue());

      final XNumericField dlgAyatToNumericField =
          DialogHelper.getNumericField(dlgDialog, DIALOG_AYAT_TO_NUMERIC_FIELD_ID);
      dlgAyatToNumericField.setValue(QuranReader.getSurahSize(selectedSurahNo));
      selectedAyatTo = (int) Math.round(dlgAyatToNumericField.getValue());
    }
  }

  private void handleArabicCheckButtonPressed() {
    final XCheckBox dlgArabicCheckbox =
        DialogHelper.getCheckBox(dlgDialog, DIALOG_ARABIC_CHECKBOX_ID);

    selectedArabicInd = short2Boolean(dlgArabicCheckbox.getState());

    DialogHelper.enableComponent(dlgDialog, DIALOG_ARABIC_LISTBOX_ID, selectedArabicInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_ARABIC_FONT_GROUPBOX_ID, selectedArabicInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_ARABIC_FONT_LABEL_ID, selectedArabicInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_ARABIC_FONT_LISTBOX_ID, selectedArabicInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_ARABIC_FONTSIZE_LABEL_ID, selectedArabicInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_ARABIC_FONTSIZE_NUMERICFIELD_ID,
        selectedArabicInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_OK_BUTTON_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_MISCELLANEOUS_GROUPBOX_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_LINE_BY_LINE_CHECKBOX_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_LINE_BY_LINE_LABEL_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
  }

  private void handleArabicFontSelected() {
    final XListBox dlgArabicFontListBox =
        DialogHelper.getListBox(dlgDialog, DIALOG_ARABIC_FONT_LISTBOX_ID);

    selectedArabicFontName = dlgArabicFontListBox.getSelectedItem();
  }

  private void handleArabicFontTextChanged() {
    final XNumericField dlgArabicFrontSize =
        DialogHelper.getNumericField(dlgDialog, DIALOG_ARABIC_FONTSIZE_NUMERICFIELD_ID);

    selectedArabicFontSize = dlgArabicFrontSize.getValue();
  }

  private void handleArabicVersionSelected() {
    final XListBox dlgArabicListBox = DialogHelper.getListBox(dlgDialog, DIALOG_ARABIC_LISTBOX_ID);

    selectedArabicLanguage =
        InsertQuranTextDialog.getItemLanguague(dlgArabicListBox.getSelectedItem());
    selectedArabicVersion =
        InsertQuranTextDialog.getItemVersion(dlgArabicListBox.getSelectedItem());
  }

  private void handleAyatFromTextChanged() {
    final XNumericField dlgAyatFromNumericField =
        DialogHelper.getNumericField(dlgDialog, DIALOG_AYAT_FROM_NUMERIC_FIELD_ID);

    if (Math.round(dlgAyatFromNumericField.getValue()) >= selectedAyatTo) {
      dlgAyatFromNumericField.setValue(selectedAyatTo);
    }

    selectedAyatFrom = (int) Math.round(dlgAyatFromNumericField.getValue());
  }

  private void handleAyatToTextChanged() {
    final XNumericField dlgAyatToNumericField =
        DialogHelper.getNumericField(dlgDialog, DIALOG_AYAT_TO_NUMERIC_FIELD_ID);

    if (Math.round(dlgAyatToNumericField.getValue()) <= selectedAyatFrom) {
      dlgAyatToNumericField.setValue(selectedAyatFrom);
    } else if (Math.round(dlgAyatToNumericField.getValue()) >= QuranReader
        .getSurahSize(selectedSurahNo)) {
      dlgAyatToNumericField.setValue(QuranReader.getSurahSize(selectedSurahNo));
    }

    selectedAyatTo = (int) Math.round(dlgAyatToNumericField.getValue());
  }

  private void handleLineByLineCheckButtonPressed() {
    final XCheckBox dlgLineByLineCheckbox =
        DialogHelper.getCheckBox(dlgDialog, DIALOG_LINE_BY_LINE_CHECKBOX_ID);

    selectedLineByLineInd = short2Boolean(dlgLineByLineCheckbox.getState());
  }

  private void handleNonArabicFontSelected() {
    final XListBox dlgNonArabicFontListBox =
        DialogHelper.getListBox(dlgDialog, DIALOG_NON_ARABIC_FONT_LISTBOX_ID);

    selectedNonArabicFontName = dlgNonArabicFontListBox.getSelectedItem();
  }

  private void handleNonArabicFontTextChanged() {
    final XNumericField dlgNonArabicFrontSize =
        DialogHelper.getNumericField(dlgDialog, DIALOG_NON_ARABIC_FONTSIZE_NUMERICFIELD_ID);

    selectedNonArabicFontSize = dlgNonArabicFrontSize.getValue();
  }

  /**
   * Handler for Ok Button.
   */
  private void handleOkButtonPressed() {
    writeSurah(selectedSurahNo);
    dlgDialog.endExecute();
  }

  private void handleSurahSelected() {
    final XListBox dlgSurahListBox = DialogHelper.getListBox(dlgDialog, DIALOG_SURAH_LISTBOX_ID);

    selectedSurahNo = dlgSurahListBox.getSelectedItemPos() + 1;

    final XNumericField dlgAyatFromNumericField =
        DialogHelper.getNumericField(dlgDialog, DIALOG_AYAT_FROM_NUMERIC_FIELD_ID);
    dlgAyatFromNumericField.setValue(1);
    selectedAyatFrom = (int) Math.round(dlgAyatFromNumericField.getValue());

    final XNumericField dlgAyatToNumericField =
        DialogHelper.getNumericField(dlgDialog, DIALOG_AYAT_TO_NUMERIC_FIELD_ID);
    dlgAyatToNumericField.setValue(QuranReader.getSurahSize(selectedSurahNo));
    selectedAyatTo = (int) Math.round(dlgAyatToNumericField.getValue());
  }

  private void handleTranslationCheckButtonPressed() {
    final XCheckBox dlgTranslationCheckbox =
        DialogHelper.getCheckBox(dlgDialog, DIALOG_TRANSLATION_CHECKBOX_ID);

    selectedTranslationInd = short2Boolean(dlgTranslationCheckbox.getState());

    DialogHelper.enableComponent(dlgDialog, DIALOG_TRANSLATION_LISTBOX_ID, selectedTranslationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_NON_ARABIC_FONT_GROUPBOX_ID,
        selectedTransliterationInd || selectedTranslationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_NON_ARABIC_FONT_LABEL_ID,
        selectedTransliterationInd || selectedTranslationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_NON_ARABIC_FONT_LISTBOX_ID,
        selectedTransliterationInd || selectedTranslationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_NON_ARABIC_FONTSIZE_LABEL_ID,
        selectedTransliterationInd || selectedTranslationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_NON_ARABIC_FONTSIZE_NUMERICFIELD_ID,
        selectedTransliterationInd || selectedTranslationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_OK_BUTTON_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_MISCELLANEOUS_GROUPBOX_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_LINE_BY_LINE_CHECKBOX_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_LINE_BY_LINE_LABEL_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
  }

  private void handleTranslationVersionSelected() {
    final XListBox dlgTranslationListBox =
        DialogHelper.getListBox(dlgDialog, DIALOG_TRANSLATION_LISTBOX_ID);

    selectedTranslationLanguage =
        InsertQuranTextDialog.getItemLanguague(dlgTranslationListBox.getSelectedItem());
    selectedTranslationVersion =
        InsertQuranTextDialog.getItemVersion(dlgTranslationListBox.getSelectedItem());
  }

  private void handleTransliterationCheckButtonPressedSelected() {
    final XCheckBox dlgTransliterationCheckbox =
        DialogHelper.getCheckBox(dlgDialog, DIALOG_TRANSLITERATION_CHECKBOX_ID);

    selectedTransliterationInd = short2Boolean(dlgTransliterationCheckbox.getState());

    DialogHelper.enableComponent(dlgDialog, DIALOG_TRANSLITERATION_LISTBOX_ID,
        selectedTransliterationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_NON_ARABIC_FONT_GROUPBOX_ID,
        selectedTransliterationInd || selectedTranslationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_NON_ARABIC_FONT_LABEL_ID,
        selectedTransliterationInd || selectedTranslationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_NON_ARABIC_FONT_LISTBOX_ID,
        selectedTransliterationInd || selectedTranslationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_NON_ARABIC_FONTSIZE_LABEL_ID,
        selectedTransliterationInd || selectedTranslationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_NON_ARABIC_FONTSIZE_NUMERICFIELD_ID,
        selectedTransliterationInd || selectedTranslationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_OK_BUTTON_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_MISCELLANEOUS_GROUPBOX_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_LINE_BY_LINE_CHECKBOX_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
    DialogHelper.enableComponent(dlgDialog, DIALOG_LINE_BY_LINE_LABEL_ID,
        selectedArabicInd || selectedTranslationInd || selectedTransliterationInd);
  }

  private void handleTransliterationVersionSelected() {
    final XListBox dlgTransliterationListBox =
        DialogHelper.getListBox(dlgDialog, DIALOG_TRANSLITERATION_LISTBOX_ID);

    selectedTransliterationLanguage =
        InsertQuranTextDialog.getItemLanguague(dlgTransliterationListBox.getSelectedItem());
    selectedTransliterationVersion =
        InsertQuranTextDialog.getItemVersion(dlgTransliterationListBox.getSelectedItem());
  }

  /**
   * Iniializes the all ayat checkbox.
   */
  private void initializeAllAyatCheckBox() {
    final XCheckBox dlgAllAyatCheckbox =
        DialogHelper.getCheckBox(dlgDialog, DIALOG_ALL_AYAT_CHECKBOX_ID);
    dlgAllAyatCheckbox.setState(InsertQuranTextDialog.boolean2Short(true));
    selectedAllAyatInd = InsertQuranTextDialog.short2Boolean(dlgAllAyatCheckbox.getState());
  }

  /**
   * Iniializes the Arabic checkbox.
   */
  private void initializeArabicCheckBox() {
    final XCheckBox dlgArabicCheckbox =
        DialogHelper.getCheckBox(dlgDialog, DIALOG_ARABIC_CHECKBOX_ID);
    dlgArabicCheckbox.setState(InsertQuranTextDialog.boolean2Short(true));
    selectedArabicInd = InsertQuranTextDialog.short2Boolean(dlgArabicCheckbox.getState());
  }

  /**
   * Iniializes the listbox with all the fonts that support Arabic characters.
   */
  private void initializeArabicFontListBox() {
    final XListBox dlgArabicFontListBox =
        DialogHelper.getListBox(dlgDialog, DIALOG_ARABIC_FONT_LISTBOX_ID);

    final Locale locale = new Locale.Builder().setScript("ARAB").build();
    final String[] fonts =
        GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames(locale);
    for (int i = 0; i < fonts.length; i++) {

      if (new Font(fonts[i], Font.PLAIN, 10).canDisplay(0x0627)) { // If Alif -> Arabic support
        dlgArabicFontListBox.addItem(fonts[i], (short) i);
        if (fonts[i].equals(getDefaultArabicFontName())) {
          dlgArabicFontListBox.selectItemPos((short) i, true);
        }
      }
    }
    dlgArabicFontListBox.selectItem(getDefaultArabicFontName(), true);
    selectedArabicFontName = dlgArabicFontListBox.getSelectedItem();
  }

  private void initializeArabicFontSize() {
    final XNumericField dlgArabicFontSizeNumericField =
        DialogHelper.getNumericField(dlgDialog, DIALOG_ARABIC_FONTSIZE_NUMERICFIELD_ID);

    dlgArabicFontSizeNumericField.setValue(getDefaultArabicCharHeight());
    selectedArabicFontSize = getDefaultArabicCharHeight();
  }

  private void initializeArabicLineNumberCheckBox() {
    selectedLineNumberInd = true;
  }

  /**
   * Iniializes the listbox with all the Arabic versions of the Qur'an.
   */
  private void initializeArabicListBox() {
    final XListBox dlgArabicListBox = DialogHelper.getListBox(dlgDialog, DIALOG_ARABIC_LISTBOX_ID);

    QuranReader.getAllQuranVersions().entrySet().stream().forEach(entry -> {
      final String currentKey = entry.getKey();
      final String[] currentValue = entry.getValue();
      int i = 0;
      for (final String element : currentValue) {
        if (currentKey.equals(ARABIC)) {
          dlgArabicListBox.addItem(currentKey + " (" + element + ")", (short) i++);
        }
      }
    });
    if (dlgArabicListBox.getItemCount() > 0) {
      dlgArabicListBox.selectItemPos((short) 0, true);
      selectedArabicLanguage =
          InsertQuranTextDialog.getItemLanguague(dlgArabicListBox.getSelectedItem());
      selectedArabicVersion =
          InsertQuranTextDialog.getItemVersion(dlgArabicListBox.getSelectedItem());
    }
  }

  private void initializeAyatFrom() {
    final XNumericField dlgAyatFromNumericField =
        DialogHelper.getNumericField(dlgDialog, DIALOG_AYAT_FROM_NUMERIC_FIELD_ID);

    dlgAyatFromNumericField.setValue(1);
  }

  private void initializeAyatTo() {
    final XNumericField dlgAyatToNumericField =
        DialogHelper.getNumericField(dlgDialog, DIALOG_AYAT_TO_NUMERIC_FIELD_ID);

    dlgAyatToNumericField.setValue(QuranReader.getSurahSize(selectedSurahNo));
  }

  /**
   * Initializes the dialog.
   */
  private void initializeDialog() {
    getLoDocumentDefaults();
    initializeSurahListBox();
    initializeAllAyatCheckBox();
    initializeAyatTo();
    initializeAyatFrom();
    initializeArabicCheckBox();
    initializeArabicListBox();
    initializeTranslationCheckBox();
    initializeTranslationListBox();
    initializeTransliterationCheckBox();
    initializeTransliterationListBox();
    initializeArabicFontListBox();
    initializeArabicFontSize();
    initializeArabicLineNumberCheckBox();
    initializeNonArabicFontListBox();
    initializeNonArabicFontSize();
    initializeLineByLineCheckBox();
    initializeWriteSurahProgressBar();
  }

  /**
   * Iniializes the Arabic checkbox.
   */
  private void initializeLineByLineCheckBox() {
    final XCheckBox dlgLineByLineCheckbox =
        DialogHelper.getCheckBox(dlgDialog, DIALOG_LINE_BY_LINE_CHECKBOX_ID);
    dlgLineByLineCheckbox.setState(InsertQuranTextDialog.boolean2Short(true));
    selectedLineByLineInd = InsertQuranTextDialog.short2Boolean(dlgLineByLineCheckbox.getState());
  }

  /**
   * Initializes the listbox with all the fonts that support Latin characters.
   */
  private void initializeNonArabicFontListBox() {
    final XListBox dlgNonArabicFontListBox =
        DialogHelper.getListBox(dlgDialog, DIALOG_NON_ARABIC_FONT_LISTBOX_ID);

    final Locale locale = new Locale.Builder().setScript("LATN").build();
    final String[] fonts =
        GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames(locale);
    for (int i = 0; i < fonts.length; i++) {
      dlgNonArabicFontListBox.addItem(fonts[i], (short) i);
      if (fonts[i].equals(getDefaultArabicFontName())) {
        dlgNonArabicFontListBox.selectItemPos((short) i, true);
      }
    }
    dlgNonArabicFontListBox.selectItem(getDefaultNonArabicFontName(), true);
    selectedNonArabicFontName = dlgNonArabicFontListBox.getSelectedItem();
  }

  private void initializeNonArabicFontSize() {
    final XNumericField dlgNonArabicFontSizeNumericField =
        DialogHelper.getNumericField(dlgDialog, DIALOG_NON_ARABIC_FONTSIZE_NUMERICFIELD_ID);

    dlgNonArabicFontSizeNumericField.setValue(getDefaultNonArabicCharHeight());
    selectedNonArabicFontSize = getDefaultNonArabicCharHeight();
  }

  /**
   * Initializes the listbox with all the surah names of the Qur'an.
   */
  private void initializeSurahListBox() {
    final XListBox dlgSurahListBox = DialogHelper.getListBox(dlgDialog, DIALOG_SURAH_LISTBOX_ID);

    for (int i = 0; i < 114; i++) {
      dlgSurahListBox.addItem(QuranReader.getSurahName(i + 1) + " (" + (i + 1) + ")", (short) i);
    }
    dlgSurahListBox.selectItemPos((short) 0, true);
    selectedSurahNo = dlgSurahListBox.getSelectedItemPos() + 1;
  }

  /**
   * Iniializes the Translation checkbox.
   */
  private void initializeTranslationCheckBox() {
    final XCheckBox dlgTranslationCheckbox =
        DialogHelper.getCheckBox(dlgDialog, DIALOG_TRANSLATION_CHECKBOX_ID);
    dlgTranslationCheckbox.setState(InsertQuranTextDialog.boolean2Short(false));
    selectedTranslationInd = InsertQuranTextDialog.short2Boolean(dlgTranslationCheckbox.getState());
  }

  /**
   * Iniializes the listbox with all the Translation versions of the Qur'an.
   */
  private void initializeTranslationListBox() {
    final XListBox dlgTranslationListBox =
        DialogHelper.getListBox(dlgDialog, DIALOG_TRANSLATION_LISTBOX_ID);

    QuranReader.getAllQuranVersions().entrySet().stream().forEach(entry -> {
      final String currentKey = entry.getKey();
      final String[] currentValue = entry.getValue();
      int i = 0;
      for (final String element : currentValue) {
        if (!currentKey.equals(ARABIC)) {
          dlgTranslationListBox.addItem(currentKey + " (" + element + ")", (short) i++);
        }
      }
    });
    if (dlgTranslationListBox.getItemCount() > 0) {
      dlgTranslationListBox.selectItemPos((short) 0, true);
      selectedTranslationLanguage =
          InsertQuranTextDialog.getItemLanguague(dlgTranslationListBox.getSelectedItem());
      selectedTranslationVersion =
          InsertQuranTextDialog.getItemVersion(dlgTranslationListBox.getSelectedItem());
    }
  }

  /**
   * Iniializes the transliteration checkbox.
   */
  private void initializeTransliterationCheckBox() {
    final XCheckBox dlgTransliterationCheckbox =
        DialogHelper.getCheckBox(dlgDialog, DIALOG_TRANSLITERATION_CHECKBOX_ID);
    dlgTransliterationCheckbox.setState(InsertQuranTextDialog.boolean2Short(false));
    selectedTransliterationInd =
        InsertQuranTextDialog.short2Boolean(dlgTransliterationCheckbox.getState());
  }



  /**
   * Iniializes the listbox with all the Transliteration versions of the Qur'an.
   */
  private void initializeTransliterationListBox() {
    final XListBox dlgTransliterationListBox =
        DialogHelper.getListBox(dlgDialog, DIALOG_TRANSLITERATION_LISTBOX_ID);

    QuranReader.getAllQuranVersions().entrySet().stream().forEach(entry -> {
      final String currentKey = entry.getKey();
      final String[] currentValue = entry.getValue();
      int i = 0;
      for (final String element : currentValue) {
        if (currentKey.equals("Transliteration")) {
          dlgTransliterationListBox.addItem(currentKey + " (" + element + ")", (short) i++);
        }
      }
    });
    if (dlgTransliterationListBox.getItemCount() > 0) {
      dlgTransliterationListBox.selectItemPos((short) 0, true);
      selectedTransliterationLanguage =
          InsertQuranTextDialog.getItemLanguague(dlgTransliterationListBox.getSelectedItem());
      selectedTransliterationVersion =
          InsertQuranTextDialog.getItemVersion(dlgTransliterationListBox.getSelectedItem());
    }
  }

  private void initializeWriteSurahProgressBar() {
    final XProgressBar dlgWriteSurahProgressBar =
        DialogHelper.getProgressBar(dlgDialog, DIALOG_WRITE_SURAH_PROGRESSBAR);

    dlgWriteSurahProgressBar.setRange(0, 100);
    dlgWriteSurahProgressBar.setValue(0);
  }

  /**
   * Show the dialog.
   */
  public void show() {
    initializeDialog();
    dlgDialog.execute();
  }

  private void writeParagraph(final XText text, final XParagraphCursor paragraphCursor,
      final String paragraph, final ParagraphAdjust alignment, final short writingMode)
      throws UnknownPropertyException, PropertyVetoException, WrappedTargetException {

    paragraphCursor.gotoEndOfParagraph(false);
    text.insertControlCharacter(paragraphCursor, ControlCharacter.PARAGRAPH_BREAK, false);

    final XPropertySet paragraphCursorPropertySet = DocumentHelper.getPropertySet(paragraphCursor);
    paragraphCursorPropertySet.setPropertyValue("ParaAdjust", alignment);
    paragraphCursorPropertySet.setPropertyValue("WritingMode", writingMode);
    text.insertString(paragraphCursor, paragraph, false);
  }

  /**
   * Write the selected surah text.
   *
   * @param surahNumber the surah number
   */
  public void writeSurah(final int surahNumber) {
    final XTextDocument textDoc = DocumentHelper.getCurrentDocument(dlgContext);
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
      paragraphCursorPropertySet.setPropertyValue("CharFontName", selectedNonArabicFontName);
      paragraphCursorPropertySet.setPropertyValue("CharFontNameComplex", selectedArabicFontName);
      paragraphCursorPropertySet.setPropertyValue("CharHeight", selectedNonArabicFontSize);
      paragraphCursorPropertySet.setPropertyValue("CharHeightComplex", selectedArabicFontSize);

      final int from = (selectedAllAyatInd) ? 1 : selectedAyatFrom;
      final int to =
          (selectedAllAyatInd) ? QuranReader.getSurahSize(surahNumber) + 1 : selectedAyatTo + 1;

      if (selectedLineByLineInd) {
        writeSurahLineByLine(surahNumber, text, paragraphCursor, from, to);
      } else {
        writeSurahAsOneBlock(surahNumber, text, paragraphCursor, from, to);
      }
    } catch (com.sun.star.lang.IllegalArgumentException | UnknownPropertyException
        | PropertyVetoException | WrappedTargetException e) {
      e.printStackTrace();
    }
  }

  /**
   * Write the selected Surah as one block.
   *
   * @param surahNumber the surah number
   * @param text the text
   * @param paragraphCursor the paragraph
   * @param from start ayat
   * @param to last ayat
   */
  private void writeSurahAsOneBlock(final int surahNumber, final XText text,
      final XParagraphCursor paragraphCursor, final int from, final int to) {
    try {
      final XProgressBar dlgWriteSurahProgressBar =
          DialogHelper.getProgressBar(dlgDialog, DIALOG_WRITE_SURAH_PROGRESSBAR);
      if (selectedArabicInd) {
        StringBuilder la = new StringBuilder();
        for (int l = from; l < to; l++) {
          dlgWriteSurahProgressBar.setValue(100 * l / (to - from + 1));
          la.append(getAyahLine(surahNumber, l, selectedArabicLanguage, selectedArabicVersion));
          la.append(" ");
        }
        writeParagraph(text, paragraphCursor, la.toString() + "\n",
            com.sun.star.style.ParagraphAdjust.RIGHT, com.sun.star.text.WritingMode2.RL_TB);
      }
      if (selectedTranslationInd) {
        StringBuilder lb = new StringBuilder();
        for (int l = from; l < to; l++) {
          dlgWriteSurahProgressBar.setValue(100 * l / (to - from + 1));
          lb.append(
              getAyahLine(surahNumber, l, selectedTranslationLanguage, selectedTranslationVersion));
          lb.append(" ");
        }
        writeParagraph(text, paragraphCursor, lb.toString() + "\n",
            com.sun.star.style.ParagraphAdjust.LEFT, com.sun.star.text.WritingMode2.LR_TB);
      }
      if (selectedTransliterationInd) {
        StringBuilder lc = new StringBuilder();
        for (int l = from; l < to; l++) {
          dlgWriteSurahProgressBar.setValue(100 * l / (to - from + 1));
          lc.append(getAyahLine(surahNumber, l, selectedTransliterationLanguage,
              selectedTransliterationVersion));
          lc.append(" ");
        }
        writeParagraph(text, paragraphCursor, lc.toString() + "\n",
            com.sun.star.style.ParagraphAdjust.LEFT, com.sun.star.text.WritingMode2.LR_TB);
      }
    } catch (com.sun.star.lang.IllegalArgumentException | UnknownPropertyException
        | PropertyVetoException | WrappedTargetException e) {
      e.printStackTrace();
    }
  }

  /**
   * Write the selected Surah Line By Line.
   *
   * @param surahNumber the surah number
   * @param text the text
   * @param paragraphCursor the paragraph of the text
   * @param from start ayat
   * @param to last ayat
   */
  private void writeSurahLineByLine(final int surahNumber, final XText text,
      final XParagraphCursor paragraphCursor, final int from, final int to) {
    try {
      final XProgressBar dlgWriteSurahProgressBar =
          DialogHelper.getProgressBar(dlgDialog, DIALOG_WRITE_SURAH_PROGRESSBAR);
      for (int l = from; l < to; l++) {
        dlgWriteSurahProgressBar.setValue(100 * l / (to - from + 1));
        if (selectedArabicInd) {
          writeParagraph(text, paragraphCursor,
              getAyahLine(surahNumber, l, selectedArabicLanguage, selectedArabicVersion) + "\n",
              com.sun.star.style.ParagraphAdjust.RIGHT, com.sun.star.text.WritingMode2.RL_TB);
        }
        if (selectedTranslationInd) {
          writeParagraph(text, paragraphCursor,
              getAyahLine(surahNumber, l, selectedTranslationLanguage, selectedTranslationVersion)
                  + "\n",
              com.sun.star.style.ParagraphAdjust.LEFT, com.sun.star.text.WritingMode2.LR_TB);
        }
        if (selectedTransliterationInd) {
          writeParagraph(text, paragraphCursor,
              getAyahLine(surahNumber, l, selectedTransliterationLanguage,
                  selectedTransliterationVersion) + "\n",
              com.sun.star.style.ParagraphAdjust.LEFT, com.sun.star.text.WritingMode2.LR_TB);
        }
      }
    } catch (com.sun.star.lang.IllegalArgumentException | UnknownPropertyException
        | PropertyVetoException | WrappedTargetException e) {
      e.printStackTrace();
    }
  }

}

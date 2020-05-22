/*
 * This file is part of QuranLO
 * 
 * Copyright (C) 2020  <mossie@mossoft.nl>
 * 
 * QuranLO is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

package nl.mossoft.loeiqt.dialog;

import java.awt.Font;
import java.awt.GraphicsEnvironment;
import java.util.List;
import java.util.Locale;

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
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextViewCursor;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import nl.mossoft.loeiqt.helper.DialogHelper;
import nl.mossoft.loeiqt.helper.DocumentHelper;
import nl.mossoft.loeiqt.helper.QuranReader;

public class InsertQuranTextDialog implements XDialogEventHandler {

    private static final String EVENTALLRANGEINDICATORCHANGED = "EvtAllRangeIndicatorChanged";
    private static final String EVENTAYAHFROMCHANGED = "EvtAyahFromChanged";
    private static final String EVENTAYAHTOCHANGED = "EvtAyahToChanged";
    private static final String EVENTFONTNAMECHANGED = "EvtFontNameChanged";
    private static final String EVENTFONTSIZECHANGED = "EvtFontSizeChanged";
    private static final String EVENTNEWLINEINDICATORCHANGED = "EvtNewLineIndicatorChanged";
    private static final String EVENTOKBUTTONPRESSED = "EvtOkButtonPressed";
    private static final String EVENTSURAHCHANGED = "EvtSurahChanged";

    private static short booleanToShort(boolean b) {
	return (short) (b ? 1 : 0);
    }

    private static String numToArabNum(int n) {
	String[] arabNum = { "\u0660", "\u0661", "\u0662", "\u0663", "\u0664", "\u0665", "\u0666", "\u0667", "\u0668",
		"\u0669" };

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

    private String defaultCharFontNameComplex;
    private float defaultCharHeightComplex;

    private XCheckBox dlgAllRangeCheckBox;
    private XNumericField dlgAyatFromNumericField;
    private XNumericField dlgAyatToNumericField;
    private XComponentContext dlgContext;
    private XDialog dlgDialog;
    private XListBox dlgFontListBox;
    private XNumericField dlgFontSizeNumericField;
    private XCheckBox dlgNewlineCheckBox;

    private XListBox dlgSurahListBox;
    private boolean selectedAllRangeIndicator;
    private int selectedAyatFrom;
    private int selectedAyatTo;
    private String selectedFontNameComplex;
    private double selectedFontSize;
    private boolean selectedNewlineIndicator;

    private short selectedSurahNum;
    private String[] supportedActions = new String[] { EVENTALLRANGEINDICATORCHANGED, EVENTAYAHFROMCHANGED,
	    EVENTAYAHTOCHANGED, EVENTOKBUTTONPRESSED, EVENTNEWLINEINDICATORCHANGED, EVENTSURAHCHANGED,
	    EVENTFONTSIZECHANGED, EVENTFONTNAMECHANGED };

    public InsertQuranTextDialog(XComponentContext xContext) {
	dlgDialog = DialogHelper.createDialog("InsertQuranTextDialog.xdl", xContext, this);
	dlgContext = xContext;
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
	} else if (methodName.equals(EVENTNEWLINEINDICATORCHANGED)) {
	    dlgHandlerNewLineCheckbox();
	    return true;
	} else if (methodName.equals(EVENTFONTNAMECHANGED)) {
	    dlgHandlerFontListBox();
	    return true;
	} else if (methodName.equals(EVENTFONTSIZECHANGED)) {
	    dlgHandlerFontSizeNumericField();
	    return true;
	} else if (methodName.equals(EVENTSURAHCHANGED)) {
	    dlgHandlerSurahListBox();
	    return true;
	}

	return false;
    }

    private void configureDialogOnOpening() {
	getLODocumentDefaults();
	initializeSurahListBox();
	initializeAyatGroupBox();
	initializeAllRangeCheckBox();
	initializeFontListBox();
	initializeFontSizeNumericField();
	initializeNewlineCheckBox();
    }

    private void createOutput() {
	XTextDocument xTextDoc = DocumentHelper.getCurrentDocument(dlgContext);
	XController xController = xTextDoc.getCurrentController();
	XTextViewCursorSupplier xTextViewCursorSupplier = DocumentHelper.getCursorSupplier(xController);
	XTextViewCursor xTextViewCursor = xTextViewCursorSupplier.getViewCursor();
	XText xText = xTextViewCursor.getText();
	XTextCursor xTextCursor = xText.createTextCursorByRange(xTextViewCursor.getStart());
	XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
	XPropertySet xParagraphCursorPropertySet = DocumentHelper.getPropertySet(xParagraphCursor);

	try {
	    xParagraphCursorPropertySet.setPropertyValue("WritingMode", com.sun.star.text.WritingMode2.RL_TB);
	    xParagraphCursorPropertySet.setPropertyValue("ParaAdjust", com.sun.star.style.ParagraphAdjust.RIGHT);
	    xParagraphCursorPropertySet.setPropertyValue("CharHeightComplex", selectedFontSize);
	    xParagraphCursorPropertySet.setPropertyValue("CharFontNameComplex", selectedFontNameComplex);

	    QuranReader qrReader = new QuranReader("ar", "1", dlgContext);
	    if (selectedAllRangeIndicator) {
		if (selectedSurahNum != 1 && selectedSurahNum != 9) {
		    createOutputAyah(QuranReader.getBismillah(), 0, xParagraphCursor);
		}
		List<String> ayat = qrReader.getAllAyatOfSuraNo(selectedSurahNum);
		for (int i = 0; i < ayat.size(); i++) {
		    createOutputAyah(ayat.get(i), i + 1, xParagraphCursor);
		}

	    } else {
		List<String> ayat = qrReader.getAyatFromToOfSuraNo(selectedSurahNum, selectedAyatFrom, selectedAyatTo);
		for (int i = 0; i < ayat.size(); i++) {
		    createOutputAyah(ayat.get(i), i + selectedAyatFrom, xParagraphCursor);
		}
	    }
	} catch (IllegalArgumentException | UnknownPropertyException | PropertyVetoException
		| WrappedTargetException e) {
	    e.printStackTrace();
	}
    }

    private void createOutputAyah(String ayah, int ayahNo, XParagraphCursor xParagraphCursor) {
	String leftParenthesis = "\uFD3E";
	String rightParenthesis = "\uFD3F";

	xParagraphCursor.gotoEnd(false);
	xParagraphCursor.setString(ayah);

	if (ayahNo != 0) {
	    xParagraphCursor.gotoEnd(false);
	    xParagraphCursor.setString(" " + rightParenthesis + numToArabNum(ayahNo) + leftParenthesis);
	}

	xParagraphCursor.gotoEnd(false);
	xParagraphCursor.setString(selectedNewlineIndicator ? "\n" : " ");

    }

    private void dlgHandlerAllRangeCheckBox() {
	selectedAllRangeIndicator = shortToBoolean(dlgAllRangeCheckBox.getState());

	enableAyatGroupBox(!selectedAllRangeIndicator);
    }

    private void dlgHandlerAyahFromNumericField() {
	selectedAyatFrom = (int) dlgAyatFromNumericField.getValue();
    }

    private void dlgHandlerAyahToNumericField() {
	selectedAyatTo = (int) dlgAyatToNumericField.getValue();

    }

    private void dlgHandlerFontListBox() {
	selectedFontNameComplex = dlgFontListBox.getSelectedItem();
    }

    private void dlgHandlerFontSizeNumericField() {
	selectedFontSize = dlgFontSizeNumericField.getValue();
    }

    private void dlgHandlerNewLineCheckbox() {
	selectedNewlineIndicator = shortToBoolean(dlgNewlineCheckBox.getState());

    }

    private void dlgHandlerOkButton() {
	createOutput();
	dlgDialog.endExecute();
    }

    private void dlgHandlerSurahListBox() {
	selectedSurahNum = (short) (dlgSurahListBox.getSelectedItemPos() + 1);
    }

    private void enableAyatGroupBox(boolean enable) {
	try {
	    DialogHelper.enableComponent(dlgDialog, "AyatFromLabel", enable);
	    DialogHelper.enableComponent(dlgDialog, "AyatFromNumericField", enable);
	    DialogHelper.enableComponent(dlgDialog, "AyatToLabel", enable);
	    DialogHelper.enableComponent(dlgDialog, "AyatToNumericField", enable);
	    DialogHelper.enableComponent(dlgDialog, "AyatGroupBox", enable);

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
	} catch (com.sun.star.lang.IllegalArgumentException | UnknownPropertyException | PropertyVetoException
		| WrappedTargetException e) {
	    e.printStackTrace();
	}

    }

    private String getDefaultCharFontNameComplex() {
	return defaultCharFontNameComplex;
    }

    private float getDefaultCharHeightComplex() {
	return defaultCharHeightComplex;
    }

    private void getLODocumentDefaults() {
	XTextDocument xTextDoc = DocumentHelper.getCurrentDocument(dlgContext);
	XController xController = xTextDoc.getCurrentController();
	XTextViewCursorSupplier xTextViewCursorSupplier = DocumentHelper.getCursorSupplier(xController);
	XTextViewCursor xTextViewCursor = xTextViewCursorSupplier.getViewCursor();
	XText xText = xTextViewCursor.getText();
	XTextCursor xTextCursor = xText.createTextCursorByRange(xTextViewCursor.getStart());
	XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
	XPropertySet xParagraphCursorPropertySet = DocumentHelper.getPropertySet(xParagraphCursor);

	try {
	    defaultCharHeightComplex = (float) xParagraphCursorPropertySet.getPropertyValue("CharHeightComplex");
	} catch (UnknownPropertyException | WrappedTargetException e) {
	    defaultCharHeightComplex = 10;
	}
	try {
	    defaultCharFontNameComplex = (String) xParagraphCursorPropertySet.getPropertyValue("CharFontNameComplex");
	} catch (UnknownPropertyException | WrappedTargetException e) {
	    defaultCharFontNameComplex = "No Default set";
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

    private void initializeAyatGroupBox() {
	dlgAyatFromNumericField = DialogHelper.getNumericField(dlgDialog, "AyatFromNumericField");
	dlgAyatToNumericField = DialogHelper.getNumericField(dlgDialog, "AyatToNumericField");
    }

    private void initializeFontListBox() {
	final char ALIF = '\u0627';

	dlgFontListBox = DialogHelper.getListBox(dlgDialog, "FontListBox");
	Locale locale = new Locale.Builder().setScript("ARAB").build();
	String fonts[] = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames(locale);
	for (int i = 0; i < fonts.length; i++) {
	    if (new Font(fonts[i], Font.PLAIN, 10).canDisplay(ALIF)) {
		dlgFontListBox.addItem(fonts[i], (short) i);
	    }
	}
	selectedFontNameComplex = getDefaultCharFontNameComplex();

	dlgFontListBox.selectItem(selectedFontNameComplex, true);
    }

    private void initializeFontSizeNumericField() {
	dlgFontSizeNumericField = DialogHelper.getNumericField(dlgDialog, "FontSizeNumericField");

	selectedFontSize = getDefaultCharHeightComplex();

	dlgFontSizeNumericField.setValue(selectedFontSize);
    }

    private void initializeNewlineCheckBox() {
	dlgNewlineCheckBox = DialogHelper.getCheckbox(dlgDialog, "NewlineCheckBox");

	dlgNewlineCheckBox.setState(booleanToShort(true));
	selectedNewlineIndicator = shortToBoolean(dlgNewlineCheckBox.getState());
    }

    private void initializeSurahListBox() {
	dlgSurahListBox = DialogHelper.getListBox(dlgDialog, "SurahListBox");

	for (int i = 0; i < 114; i++) {
	    dlgSurahListBox.addItem(QuranReader.getSurahName(i) + " (" + (i + 1) + ")", (short) i);
	}
	dlgSurahListBox.selectItemPos((short) 0, true);
	selectedSurahNum = (short) (dlgSurahListBox.getSelectedItemPos() + 1);
    }

    public void show() {
	configureDialogOnOpening();
	dlgDialog.execute();
    }

}

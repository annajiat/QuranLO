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

package nl.mossoft.loeiqt.helper;

import com.sun.star.awt.MessageBoxType;
import com.sun.star.awt.Point;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XComboBox;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogEventHandler;
import com.sun.star.awt.XDialogProvider2;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XMessageBox;
import com.sun.star.awt.XMessageBoxFactory;
import com.sun.star.awt.XNumericField;
import com.sun.star.awt.XProgressBar;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.ucb.XFileIdentifierConverter;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import java.io.File;

/**
 * Helper class for the dialog.
 * 
 * @author abdullah
 *
 */
public class DialogHelper {

  /**
   * Returns a URL to be used with XDialogProvider to create a dialog.
   * 
   * @param context the dialog context
   * @param dialogFile the file with the dialog definition
   * @return Returns a URL to be used with XDialogProvider to create a dialog
   */
  public static String convertToUrl(XComponentContext context, File dialogFile) {
    String url = null;
    try {
      XFileIdentifierConverter fileConverter = UnoRuntime.queryInterface(
          com.sun.star.ucb.XFileIdentifierConverter.class, context.getServiceManager()
              .createInstanceWithContext("com.sun.star.ucb.FileContentProvider", context));
      url = fileConverter.getFileURLFromSystemPath("", dialogFile.getAbsolutePath());
    } catch (com.sun.star.uno.Exception ex) {
      return null;
    }
    return url;
  }

  /**
   * Create a dialog from an XDL file.
   *
   * @param xdlFile The filename in the `dialog` folder
   * @param context the dialog context
   * @return XDialog the dialog handle
   */
  public static XDialog createDialog(String xdlFile, XComponentContext context,
      XDialogEventHandler handler) {
    Object dialogProvider;
    try {
      dialogProvider = context.getServiceManager()
          .createInstanceWithContext("com.sun.star.awt.DialogProvider2", context);
      XDialogProvider2 dialogProv =
          UnoRuntime.queryInterface(XDialogProvider2.class, dialogProvider);
      File dialogFile = FileHelper.getDialogFilePath(xdlFile, context);
      return dialogProv.createDialogWithHandler(convertToUrl(context, dialogFile), handler);
    } catch (Exception e) {
      return null;
    }
  }

  /**
   * Enable button.
   * 
   * @param dialog the dialog
   * @param componentId the component id
   * @param enabled if true component will be enabled.
   */
  public static void enableButton(XDialog dialog, String componentId, boolean enabled) {
    XControlContainer dlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    // retrieve the control that we want to disable or enable
    XControl control =
        UnoRuntime.queryInterface(XControl.class, dlgContainer.getControl(componentId));
    XPropertySet propertySet = UnoRuntime.queryInterface(XPropertySet.class, control.getModel());
    try {
      propertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));
    } catch (IllegalArgumentException | UnknownPropertyException | PropertyVetoException
        | WrappedTargetException e) {
      e.printStackTrace();
    }
  }

  /**
   * Enable a component of the dialog.
   * 
   * @param dialog the dialog
   * @param componentId the component
   * @param enabled if true then the component is enabled
   */
  public static void enableComponent(XDialog dialog, String componentId, boolean enabled) {
    XControlContainer dlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    final XControl xControl =
        UnoRuntime.queryInterface(XControl.class, dlgContainer.getControl(componentId));

    final XPropertySet xPropertySet =
        UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
    try {
      xPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enabled));
    } catch (IllegalArgumentException | UnknownPropertyException | PropertyVetoException
        | WrappedTargetException e) {
      e.printStackTrace();
    }

  }

  /**
   * Returns a button (XButton) from a dialog.
   * 
   * @param dialog the dialog
   * @param componentId the component id
   * @return a button
   */
  public static XButton getButton(XDialog dialog, String componentId) {
    XControlContainer dlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    Object control = dlgContainer.getControl(componentId);
    return UnoRuntime.queryInterface(XButton.class, control);
  }

  /**
   * Returns a Ceheckbox (XCheckBox) from a dialog.
   * 
   * @param dialog the dialog
   * @param componentId the component id
   * @return the CheckBox
   */
  public static XCheckBox getCheckBox(XDialog dialog, String componentId) {
    XControlContainer dlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    Object control = dlgContainer.getControl(componentId);
    return UnoRuntime.queryInterface(XCheckBox.class, control);
  }

  /**
   * Returns a Combo box (XComboBox) from a dialog.
   * 
   * @param dialog the dialog
   * @param componentId the component id
   * @return a combobox
   */
  public static XComboBox getCombobox(XDialog dialog, String componentId) {
    XControlContainer dlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    Object control = dlgContainer.getControl(componentId);
    return UnoRuntime.queryInterface(XComboBox.class, control);
  }

  /**
   * Returns a text field (XTextComponent) from a dialog.
   * 
   * @param dialog the dialog
   * @param componentId the component id
   * @return a edit field
   */
  public static XTextComponent getEditField(XDialog dialog, String componentId) {
    XControlContainer dlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    Object control = dlgContainer.getControl(componentId);
    return UnoRuntime.queryInterface(XTextComponent.class, control);
  }

  /** Returns a label (XFixedText) from a dialog. */
  public static XFixedText getLabel(XDialog dialog, String componentId) {
    XControlContainer dlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    Object control = dlgContainer.getControl(componentId);
    return UnoRuntime.queryInterface(XFixedText.class, control);
  }

  /**
   * Returns a List box (XListBox) from a dialog.
   * 
   * @param dialog the dialog
   * @param componentId the component id
   * @return a list box
   */
  public static XListBox getListBox(XDialog dialog, String componentId) {
    XControlContainer dlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    Object control = dlgContainer.getControl(componentId);
    return UnoRuntime.queryInterface(XListBox.class, control);
  }

  /**
   * Returns a text field (XNumericField) from a dialog.
   * 
   * @param dialog the dialog
   * @param componentId the component id
   * @return a numeric field
   */
  public static XNumericField getNumericField(XDialog dialog, String componentId) {
    XControlContainer dlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    Object control = dlgContainer.getControl(componentId);
    return UnoRuntime.queryInterface(XNumericField.class, control);
  }

  /**
   * get the position of the dialog (Point).
   * 
   * @param dialog the dialog
   * @return the position
   */
  public static Point getPosition(XDialog dialog) {
    int posX = 0;
    int posY = 0;
    XControlModel dialogModel = UnoRuntime.queryInterface(XControl.class, dialog).getModel();
    XPropertySet propSet = UnoRuntime.queryInterface(XPropertySet.class, dialogModel);
    try {
      posX = (int) propSet.getPropertyValue("PositionX");
      posY = (int) propSet.getPropertyValue("PositionY");
    } catch (UnknownPropertyException | WrappedTargetException e) {
      e.printStackTrace();
    }
    return new Point(posX, posY);
  }

  /**
   * Returns a List box (XProgressBar) from a dialog.
   * 
   * @param dialog the dialog
   * @param componentId the component id
   * @return a progressbar
   */
  public static XProgressBar getProgressBar(XDialog dialog, String componentId) {
    XControlContainer dlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    Object control = dlgContainer.getControl(componentId);
    return UnoRuntime.queryInterface(XProgressBar.class, control);
  }

  /**
   * Set the focus to an input field.
   * 
   * @param editField the input field
   */
  public static void setFocus(XTextComponent editField) {
    XWindow controlWindow = UnoRuntime.queryInterface(XWindow.class, editField);
    controlWindow.setFocus();
  }

  /**
   * set the position of the dialog.
   * 
   * @param dialog the dialog
   * @param posX the x position
   * @param posY th y position
   */
  public static void setPosition(XDialog dialog, int posX, int posY) {
    XControlModel dialogModel = UnoRuntime.queryInterface(XControl.class, dialog).getModel();
    XPropertySet propSet = UnoRuntime.queryInterface(XPropertySet.class, dialogModel);
    try {
      propSet.setPropertyValue("PositionX", posX);
      propSet.setPropertyValue("PositionY", posY);
    } catch (com.sun.star.lang.IllegalArgumentException | UnknownPropertyException
        | PropertyVetoException | WrappedTargetException e) {
      e.printStackTrace();
    }
  }

  /**
   * Shows a warning message.
   * 
   * @param context the dialog context
   * @param dialog the dialog
   * @param message the message
   */
  public static void showErrorMessage(XComponentContext context, XDialog dialog, String message) {
    showMessageBox(context, dialog, MessageBoxType.ERRORBOX, "Error:", message);
  }

  /**
   * Shows a warning message.
   * 
   * @param context the dialog context
   * @param dialog the dialog
   * @param message the message
   */
  public static void showInfoMessage(XComponentContext context, XDialog dialog, String message) {
    showMessageBox(context, dialog, MessageBoxType.INFOBOX, "Info:", message);
  }

  /**
   * Show a dialog with a message of MessageBoxType type.
   * 
   * @param context the dialog context
   * @param dialog the dialog
   * @param type type of the message (MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX)
   * @param title the title of the dialog showing the message
   * @param message the message
   */
  public static void showMessageBox(XComponentContext context, XDialog dialog, MessageBoxType type,
      String title, String message) {
    XToolkit toolkit;
    try {
      toolkit = UnoRuntime.queryInterface(XToolkit.class, context.getServiceManager()
          .createInstanceWithContext("com.sun.star.awt.Toolkit", context));
    } catch (Exception e) {
      return;
    }
    XMessageBoxFactory messageBoxFactory =
        UnoRuntime.queryInterface(XMessageBoxFactory.class, toolkit);
    XWindowPeer parentWindowPeer = UnoRuntime.queryInterface(XWindowPeer.class, dialog);
    XMessageBox messageBox = messageBoxFactory.createMessageBox(parentWindowPeer, type,
        com.sun.star.awt.MessageBoxButtons.BUTTONS_OK, title, message);
    if (messageBox == null) {
      return;
    }
    messageBox.execute();
  }

  /**
   * Shows a warning message.
   * 
   * @param context the dialog context
   * @param dialog the dialog
   * @param message the message
   */
  public static void showWarningMessage(XComponentContext context, XDialog dialog, String message) {
    showMessageBox(context, dialog, MessageBoxType.WARNINGBOX, "Warning:", message);
  }

  private DialogHelper() {}

}

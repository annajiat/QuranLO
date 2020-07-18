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
import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogEventHandler;
import com.sun.star.awt.XDialogProvider2;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XMessageBox;
import com.sun.star.awt.XMessageBoxFactory;
import com.sun.star.awt.XNumericField;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import java.io.File;

/**
 * Creates Dialogs from XDL file.
 * @author abdullah
 *
 */
public class DialogHelper {

  /**
   * Convert File path to URL. 
   * 
   * @param context dialog context
   * @param dialogFile file path of dialog definition 
   * @return a URL to be used with XDialogProvider to create a dialog */
  private static String convertToUrl(XComponentContext context, File dialogFile) {
    String surl = null;
    try {
      com.sun.star.ucb.XFileIdentifierConverter xfileConverter = UnoRuntime.queryInterface(
          com.sun.star.ucb.XFileIdentifierConverter.class, context.getServiceManager()
              .createInstanceWithContext("com.sun.star.ucb.FileContentProvider", context));
      surl = xfileConverter.getFileURLFromSystemPath("", dialogFile.getAbsolutePath());
    } catch (com.sun.star.uno.Exception ex) {
      return null;
    }
    return surl;
  }

  /**
   * @param xdlFile
   * @param context
   * @param handler
   * @return
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
   * @param dialog
   * @param componentId
   * @param enable
   * @throws UnknownPropertyException
   * @throws PropertyVetoException
   * @throws WrappedTargetException
   */
  public static void enableComponent(XDialog dialog, String componentId, boolean enable)
      throws UnknownPropertyException, PropertyVetoException, WrappedTargetException {
    XControlContainer xDlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    // retrieve the control that we want to disable or enable
    XControl xControl =
        UnoRuntime.queryInterface(XControl.class, xDlgContainer.getControl(componentId));
    XPropertySet xModelPropertySet =
        UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
    xModelPropertySet.setPropertyValue("Enabled", Boolean.valueOf(enable));
  }

  /** Returns a Check box (XCheckBox) from a dialog */
  public static XCheckBox getCheckbox(XDialog dialog, String componentId) {
    XControlContainer xDlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    Object control = xDlgContainer.getControl(componentId);
    return UnoRuntime.queryInterface(XCheckBox.class, control);
  }

  /** Returns a List box (XListBox) from a dialog */
  public static XListBox getListBox(XDialog dialog, String componentId) {
    XControlContainer xDlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    Object control = xDlgContainer.getControl(componentId);
    return UnoRuntime.queryInterface(XListBox.class, control);
  }

  /** Returns a numeric field (XNumericField) from a dialog */
  public static XNumericField getNumericField(XDialog dialog, String componentId) {
    XControlContainer xDlgContainer = UnoRuntime.queryInterface(XControlContainer.class, dialog);
    Object control = xDlgContainer.getControl(componentId);
    return UnoRuntime.queryInterface(XNumericField.class, control);
  }

  public static void showErrorMessage(XComponentContext context, XDialog dialog, String message) {
    showMessageBox(context, dialog, MessageBoxType.ERRORBOX, "Fehler", message);
  }

  public static void showInfoMessage(XComponentContext context, XDialog dialog, String message) {
    showMessageBox(context, dialog, MessageBoxType.INFOBOX, "Info", message);
  }

  private static void showMessageBox(XComponentContext context, XDialog dialog, MessageBoxType type,
      String sTitle, String sMessage) {
    XToolkit xToolkit;
    try {
      xToolkit = UnoRuntime.queryInterface(XToolkit.class, context.getServiceManager()
          .createInstanceWithContext("com.sun.star.awt.Toolkit", context));
    } catch (Exception e) {
      return;
    }
    XMessageBoxFactory xMessageBoxFactory =
        UnoRuntime.queryInterface(XMessageBoxFactory.class, xToolkit);
    XWindowPeer xParentWindowPeer = UnoRuntime.queryInterface(XWindowPeer.class, dialog);
    XMessageBox xMessageBox = xMessageBoxFactory.createMessageBox(xParentWindowPeer, type,
        com.sun.star.awt.MessageBoxButtons.BUTTONS_OK, sTitle, sMessage);
    if (xMessageBox == null) {
      return;
    }

    xMessageBox.execute();
  }

  public static void showWarningMessage(XComponentContext context, XDialog dialog, String message) {
    showMessageBox(context, dialog, MessageBoxType.WARNINGBOX, "Warnung", message);
  }

  private DialogHelper() {
    throw new IllegalStateException("Utility class");
  }

}

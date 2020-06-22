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

import com.sun.star.beans.XPropertySet;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextViewCursorSupplier;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * Helps getting desktop, components, frames, cursors and other interfaces.
 *
 */
public class DocumentHelper {

  public static XComponentLoader getComponentLoader(XDesktop xDesktop) {
    return UnoRuntime.queryInterface(com.sun.star.frame.XComponentLoader.class, xDesktop);
  }

  public static XTextContent getContent(Object oGraphic) {
    return UnoRuntime.queryInterface(com.sun.star.text.XTextContent.class, oGraphic);
  }

  /** Returns the current XComponent */
  private static XComponent getCurrentComponent(XComponentContext xContext) {
    return getCurrentDesktop(xContext).getCurrentComponent();
  }

  /** Returns the current XDesktop */
  private static XDesktop getCurrentDesktop(XComponentContext xContext) {
    XMultiComponentFactory xMCF =
        UnoRuntime.queryInterface(XMultiComponentFactory.class, xContext.getServiceManager());
    Object desktop = null;
    try {
      desktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
    } catch (Exception e) {
      return null;
    }
    return UnoRuntime.queryInterface(com.sun.star.frame.XDesktop.class, desktop);
  }

  /** Returns the current text document (if any) */
  public static XTextDocument getCurrentDocument(XComponentContext xContext) {
    return UnoRuntime.queryInterface(XTextDocument.class, getCurrentComponent(xContext));
  }

  public static XTextViewCursorSupplier getCursorSupplier(XController xController) {
    return UnoRuntime.queryInterface(XTextViewCursorSupplier.class, xController);
  }

  public static XPropertySet getPropertySet(Object oGraphic) {
    return UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class, oGraphic);
  }

}

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

package nl.mossoft.loeiqt.comp;

import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.uno.XComponentContext;
import nl.mossoft.loeiqt.dialog.InsertQuranTextDialog;
import nl.mossoft.loeiqt.helper.DialogHelper;

/**
 * Implementation of the core functionality of the extension.
 * 
 * @author abdullah (mossie@mossoft.nl)
 */
public final class InsertQuranTextImpl extends WeakBase
    implements com.sun.star.lang.XServiceInfo, com.sun.star.task.XJobExecutor {

  private static final String IMPLEMENTATIONNAME = InsertQuranTextImpl.class.getName();
  private static final String[] SERVICENAMES = {"nl.mossoft.loeiqt.InsertQuranText"};

  /**
   * Get a Component Factory.
   * 
   * @param implementationName name of implementation
   * @return Component Factory
   */
  public static XSingleComponentFactory __getComponentFactory(final String implementationName) {
    XSingleComponentFactory factory = null;

    if (implementationName.equals(IMPLEMENTATIONNAME)) {
      factory = Factory.createComponentFactory(InsertQuranTextImpl.class, SERVICENAMES);
    }
    return factory;
  }

  public static boolean __writeRegistryServiceInfo(XRegistryKey registryKey) {
    return Factory.writeRegistryServiceInfo(IMPLEMENTATIONNAME, SERVICENAMES, registryKey);
  }

  private final XComponentContext context;

  public InsertQuranTextImpl(XComponentContext context) {
    this.context = context;
  }

  @Override
  public String getImplementationName() {
    return IMPLEMENTATIONNAME;
  }

  @Override
  public String[] getSupportedServiceNames() {
    return SERVICENAMES;
  }

  @Override
  public boolean supportsService(String service) {
    int len = SERVICENAMES.length;

    for (int i = 0; i < len; i++) {
      if (service.equals(SERVICENAMES[i])) {
        return true;
      }
    }
    return false;
  }

  @Override
  public void trigger(String action) {
    switch (action) {
      case "actionIqt":
        InsertQuranTextDialog insertQuranTextDialog = new InsertQuranTextDialog(context);
        insertQuranTextDialog.show();
        break;
      default:
        DialogHelper.showErrorMessage(this.context, null, "Unknown action: " + action);
    }
  }
}

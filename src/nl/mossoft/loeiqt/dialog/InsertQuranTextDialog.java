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

import com.sun.star.comp.loader.FactoryHelper;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.lang.XSingleServiceFactory;
import com.sun.star.lib.uno.helper.ComponentBase;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.task.XJobExecutor;
import com.sun.star.uno.XComponentContext;

public class InsertQuranTextDialog extends ComponentBase implements XServiceInfo, XJobExecutor {

  private static final String SERVICENAME = "nl.mossoft.loeiqt.dialog.QuranTextDialog";

  public static XSingleServiceFactory getServiceFactory(String implName,
      XMultiServiceFactory xMultiServiceFactory, XRegistryKey xRegistryKey) {

    XSingleServiceFactory xSingleServiceFactory = null;
    if (implName.equals(InsertQuranTextDialog.class.getName())) {
      xSingleServiceFactory = FactoryHelper.getServiceFactory(InsertQuranTextDialog.class,
          InsertQuranTextDialog.SERVICENAME, xMultiServiceFactory, xRegistryKey);
    }
    return xSingleServiceFactory;
  }

  private QuranTextDialog dlgQuranTextDialog;

  public InsertQuranTextDialog(XComponentContext xComponentContext) {
    dlgQuranTextDialog =
        QuranTextDialog.newInstance(xComponentContext, xComponentContext.getServiceManager());
  }

  @Override
  public String getImplementationName() {
    // TODO Auto-generated method stub
    return null;
  }

  @Override
  public String[] getSupportedServiceNames() {
    String[] retValue = new String[0];
    retValue[0] = InsertQuranTextDialog.SERVICENAME;
    return retValue;
  }

  public void show() {
    dlgQuranTextDialog.execute();
  }

  @Override
  public boolean supportsService(String serviceName) {
    if (serviceName.equals(InsertQuranTextDialog.SERVICENAME))
      return true;
    return false;
  }

  @Override
  public void trigger(String arg0) {
    // TODO Auto-generated method stub

  }
}

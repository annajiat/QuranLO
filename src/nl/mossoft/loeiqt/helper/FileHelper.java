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

package nl.mossoft.loeiqt.helper;

import java.io.File;
import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;

import com.sun.star.deployment.PackageInformationProvider;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.XURLTransformer;

class FileHelper {

    private static final String DIALOG_RESOURCES = "dialog/";
    private static final String QURAN_RESOURCES = "resources/";

    /**
     * Returns a path to a dialog file
     */
    static File getDialogFilePath(String xdlFile, XComponentContext xContext) {
	return getFilePath(DIALOG_RESOURCES + xdlFile, xContext);
    }

    /**
     * Returns a file path for a file in the installed extension, or null on
     * failure.
     */
    private static File getFilePath(String file, XComponentContext xContext) {
	XPackageInformationProvider xPackageInformationProvider = PackageInformationProvider.get(xContext);
	String location = xPackageInformationProvider.getPackageLocation("nl.mossoft.loeiqt.insertqurantext");
	Object oTransformer;
	try {
	    oTransformer = xContext.getServiceManager().createInstanceWithContext("com.sun.star.util.URLTransformer",
		    xContext);
	} catch (Exception e) {
	    e.printStackTrace();
	    return null;
	}
	XURLTransformer xTransformer = (XURLTransformer) UnoRuntime.queryInterface(XURLTransformer.class, oTransformer);
	com.sun.star.util.URL[] oURL = new com.sun.star.util.URL[1];
	oURL[0] = new com.sun.star.util.URL();
	oURL[0].Complete = location + "/" + file;
	xTransformer.parseStrict(oURL);
	URL url;
	try {
	    url = new URL(oURL[0].Complete);
	} catch (MalformedURLException e1) {
	    return null;
	}
	File f;
	try {
	    f = new File(url.toURI());
	} catch (URISyntaxException e1) {
	    return null;
	}
	return f;
    }

    /**
     * Returns a path to a Quran file
     */
    static File getQuranFilePath(String xdlFile, XComponentContext xContext) {
	return getFilePath(QURAN_RESOURCES + xdlFile, xContext);
    }

}

/*
 * The Contents of this file are made available subject to the terms of either of the GNU Lesser
 * General Public License Version 2.1
 *
 * Sun Microsystems Inc., October, 2000
 *
 * GNU Lesser General Public License Version 2.1 =============================================
 * Copyright 2000 by Sun Microsystems, Inc. 901 San Antonio Road, Palo Alto, CA 94303, USA
 *
 * This library is free software; you can redistribute it and/or modify it under the terms of the
 * GNU Lesser General Public License version 2.1, as published by the Free Software Foundation.
 *
 * This library is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
 * even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License along with this library;
 * if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA
 * 02111-1307 USA
 *
 * The Initial Developer of the Original Code is: Sun Microsystems, Inc..
 *
 * Copyright: 2002 by Sun Microsystems, Inc.
 *
 * All Rights Reserved.
 *
 * Contributor(s): Cedric Bosdonnat
 *
 */

package nl.mossoft.loeiqt.comp;

import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.registry.XRegistryKey;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.LineNumberReader;
import java.lang.reflect.Method;
import java.util.ArrayList;

/**
 * Component main registration class.
 *
 * <p>
 * <strong>This class should not be modified.</strong>
 * </p>
 *
 * @author Cedric Bosdonnat aka. cedricbosdo
 *
 */
public class RegistrationHandler {

  /**
   * Get a component factory for the implementations handled by this class.
   *
   * <p>
   * This method calls all the methods of the same name from the classes listed in the
   * <code>RegistrationHandler.classes</code> file. <strong>This method should not be
   * modified.</strong>
   * </p>
   *
   * @param implementationName the name of the implementation to create.
   * @return the factory which can create the implementation.
   */
  public static XSingleComponentFactory __getComponentFactory(final String implementationName) {
    XSingleComponentFactory factory = null;

    final Class<?>[] classes = findServicesImplementationClasses();

    int i = 0;
    while (i < classes.length && factory == null) {
      final Class<?> clazz = classes[i];
      if (implementationName.equals(clazz.getCanonicalName())) {
        try {
          final Class<?>[] getTypes = new Class[] {String.class};
          final Method getFactoryMethod = clazz.getMethod("__getComponentFactory", getTypes);
          final Object o = getFactoryMethod.invoke(null, implementationName);
          factory = (XSingleComponentFactory) o;
        } catch (final Exception e) {
          e.printStackTrace();
        }
      }
      i++;
    }
    return factory;
  }

  /**
   * Writes the services implementation informations to the UNO registry.
   *
   * <p>
   * This method calls all the methods of the same name from the classes listed in the
   * <code>RegistrationHandler.classes</code> file. <strong>This method should not be
   * modified.</strong>
   * </p>
   *
   * @param registryKey the root registry key where to write the informations.
   *
   * @return <code>true</code> if the informations have been successfully written to the registry
   *         key, <code>false</code> otherwise.
   */
  public static boolean __writeRegistryServiceInfo(final XRegistryKey registryKey) {

    final Class<?>[] classes = findServicesImplementationClasses();

    boolean success = true;
    int i = 0;
    while (i < classes.length && success) {
      final Class<?> clazz = classes[i];
      try {
        final Class<?>[] writeTypes = new Class[] {XRegistryKey.class};
        final Method getFactoryMethod = clazz.getMethod("__writeRegistryServiceInfo", writeTypes);
        final Object o = getFactoryMethod.invoke(null, registryKey);
        success = success && ((Boolean) o).booleanValue();
      } catch (final Exception e) {
        success = false;
        e.printStackTrace();
      }
      i++;
    }
    return success;
  }

  /**
   * Find implementation classes for UNO services.
   *
   * @return all the UNO implementation classes.
   */
  private static Class<?>[] findServicesImplementationClasses() {

    final ArrayList<Class<?>> classes = new ArrayList<>();

    final InputStream in =
        RegistrationHandler.class.getResourceAsStream("RegistrationHandler.classes");
    final LineNumberReader reader = new LineNumberReader(new InputStreamReader(in));

    try {
      String line = reader.readLine();
      while (line != null) {
        if (!line.equals("")) {
          line = line.trim();
          try {
            final Class<?> clazz = Class.forName(line);

            final Class<?>[] writeTypes = new Class[] {XRegistryKey.class};
            final Class<?>[] getTypes = new Class[] {String.class};

            final Method writeRegMethod = clazz.getMethod("__writeRegistryServiceInfo", writeTypes);
            final Method getFactoryMethod = clazz.getMethod("__getComponentFactory", getTypes);

            if (writeRegMethod != null && getFactoryMethod != null) {
              classes.add(clazz);
            }

          } catch (final Exception e) {
            e.printStackTrace();
          }
        }
        line = reader.readLine();
      }
    } catch (final IOException e) {
      e.printStackTrace();
    } finally {
      try {
        reader.close();
        in.close();
      } catch (final Exception e) {
        e.printStackTrace();
      }
    }

    return classes.toArray(new Class[classes.size()]);
  }

  /**
   * the Constructor.
   */
  private RegistrationHandler() {

  }
}

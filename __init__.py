# -*- coding: utf-8 -*-
"""
/***************************************************************************
 CSVProvider
                                 A QGIS plugin
 Example of "faking" a data provider with PyQGIS
                             -------------------
        begin                : 2015-03-13
        copyright            : (C) 2015 by GeoApt LLC
        email                : gsherman@geoapt.com
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""
import os 
import site
import sys

site.addsitedir(os.path.join(os.path.dirname(__file__),'extlibs'))
#global sys.argv
#sys.argv = [""]
#print sys.modules['rsa'] = None

# noinspection PyPep8Naming
def classFactory(iface):  # pylint: disable=invalid-name
    """Load CSVProvider class from file CSVProvider.

    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
    #
    from .gdrive_provider import Google_Drive_Provider
    return Google_Drive_Provider(iface)

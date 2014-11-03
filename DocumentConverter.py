#!/usr/bin/env python
#
# PyODConverter (Python OpenDocument Converter) v1.0.0 - 2008-05-05
#
# This script converts a document from one office format to another by
# connecting to an OpenOffice.org instance via Python-UNO bridge.
#
# Copyright (C) 2008 Mirko Nasato <mirko@artofsolving.com>
#                    Matthew Holloway <matthew@holloway.co.nz>
#                    Alistek Ltd. (www.alistek.com) 
# Licensed under the GNU LGPL v2.1 - http://www.gnu.org/licenses/lgpl-2.1.html
# - or any later version.
#

DEFAULT_OPENOFFICE_PORT = 8100
RESOLVESTR = "uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext"

################## For CSV documents #######################
# Field Separator (1), Text Delimiter (2), Character Set (3), Number of First Line (4)
CSVFilterOptions = "59,34,76,1"
# ASCII code of field separator
# ASCII code of text delimiter
# character set, use 0 for "system character set", 76 seems to be UTF-8
# number of first line (1-based)
# Cell format codes for the different columns (optional)
############################################################

from os.path import abspath
from os.path import isfile
from os.path import splitext
import sys
import traceback
import time
import subprocess
import logging
from io import BytesIO

import uno
import unohelper
from com.sun.star.beans import PropertyValue
from com.sun.star.uno import Exception as UnoException
from com.sun.star.connection import NoConnectException, ConnectionSetupException
from com.sun.star.beans import UnknownPropertyException
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.io import XOutputStream
from com.sun.star.io import IOException


class DocumentConversionException(Exception):

    def __init__(self, message):
        self.message = message

    def __str__(self):
        return self.message

class OutputStreamWrapper(unohelper.Base, XOutputStream):
    """ Minimal Implementation of XOutputStream """
    def __init__(self, debug=True):
        self.debug = debug
        self.data = BytesIO()
        self.position = 0
        if self.debug:
            sys.stderr.write("__init__ OutputStreamWrapper.\n")

    def writeBytes(self, bytes):
        if self.debug:
            sys.stderr.write("writeBytes %i bytes.\n" % len(bytes.value))
        self.data.write(bytes.value)
        self.position += len(bytes.value)

    def close(self):
        if self.debug:
            sys.stderr.write("Closing output. %i bytes written.\n" % self.position)
        self.data.close()

    def flush(self):
        if self.debug:
            sys.stderr.write("Flushing output.\n")
        pass
    def closeOutput(self):
        if self.debug:
            sys.stderr.write("Closing output.\n")
        pass

class DocumentConverter:
   
    def __init__(self, host='localhost', port=DEFAULT_OPENOFFICE_PORT, ooo_restart_cmd=None):
        self._host = host
        self._port = port
        self.logger = logging.getLogger('main')
        self._ooo_restart_cmd = ooo_restart_cmd
        self.localContext = uno.getComponentContext()
        self.serviceManager = self.localContext.ServiceManager
        self._resolver = self.serviceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", self.localContext)
        try:
            self._context = self._resolver.resolve(RESOLVESTR % (host, port))
        except IllegalArgumentException as exception:
            raise DocumentConversionException("The url is invalid (%s)" % exception)
        except NoConnectException as exception:
            if self._restart_ooo():
                # We try again once
                try:
                    self._context = self._resolver.resolve(RESOLVESTR % (host, port))
                except NoConnectException as exception:
                    raise DocumentConversionException("Failed to connect to OpenOffice.org on host %s, port %s. %s" % (host, port, exception))
            else:
                raise DocumentConversionException("Failed to connect to OpenOffice.org on host %s, port %s. %s" % (host, port, exception))

        except ConnectionSetupException as exception:
            raise DocumentConversionException("Not possible to accept on a local resource (%s)" % exception)

    def putDocument(self, data):
        try:
            self.desktop = self._context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self._context)
        except UnknownPropertyException:
            self._context = self._resolver.resolve(RESOLVESTR % (self._host, self._port))
            self.desktop = self._context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self._context)
        inputStream = self.serviceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", self.localContext)
        inputStream.initialize((uno.ByteSequence(data),))
        props = self._toProperties(InputStream = inputStream, FilterName="writer8")
        try:
            self.document = self.desktop.loadComponentFromURL('private:stream', "_blank", 0, props)
        except:
            exceptionType, exceptionValue, exceptionTraceback = sys.exc_info()
            traceback.print_exception(
                            exceptionType,
                            exceptionValue,
                            exceptionTraceback,
                            limit=2, file=sys.stdout
                            )
        inputStream.closeInput()

    def closeDocument(self):
        self.document.close(True)

    def saveByStream(self, filter_name=None):
        try:
            self.document.refresh()
        except AttributeError: # ods document does not support refresh
            pass
        outputStream = OutputStreamWrapper(False)
        props = self._toProperties(
                        OutputStream = outputStream,
                        FilterName = filter_name,
                        FilterOptions = CSVFilterOptions
                        )
        try:
            #url = uno.systemPathToFileUrl(path) #when storing to filesystem
            self.document.storeToURL('private:stream', props)
        except Exception as exception:
            exceptionType, exceptionValue, exceptionTraceback = sys.exc_info()
            traceback.print_exception(
                            exceptionType,
                            exceptionValue,
                            exceptionTraceback,
                            limit=2, file=sys.stdout
                            )
        openDocumentBytes = outputStream.data.getvalue()
        outputStream.close()
        return openDocumentBytes
        

    def insertSubreports(self, oo_subreports):
        """
        Inserts the given file into the current document.
        The file contents will replace the placeholder text.
        """
        import os

        for subreport in oo_subreports:
            fd = file(subreport, 'rb')
            placeholder_text = "<insert_doc('%s')>" % subreport
            subdata = fd.read()
            subStream = self.serviceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", self.localContext)
            subStream.initialize((uno.ByteSequence(subdata),))

            search = self.document.createSearchDescriptor()
            search.SearchString = placeholder_text
            found = self.document.findFirst( search )
            #while found:
            props = self._toProperties(InputStream = subStream, FilterName = "writer8")
            try:
                found.insertDocumentFromURL('private:stream', props)
            except Exception as ex:
                print (_("Error inserting file %s on the OpenOffice document: %s") % (subreport, ex))
                exceptionType, exceptionValue, exceptionTraceback = sys.exc_info()
                traceback.print_exception(exceptionType, exceptionValue, exceptionTraceback,
                                      limit=2, file=sys.stdout)
            #found = self.document.findNext(found, search)

            os.unlink(subreport)

    def joinDocuments(self, docs):
        while(docs):
            subStream = self.serviceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", self.localContext)
            subStream.initialize((uno.ByteSequence(docs.pop()),))
            props = self._toProperties(InputStream = subStream, FilterName = "writer8")
            try:
                self.document.Text.getEnd().insertDocumentFromURL('private:stream', props)
            except Exception as exception:
                print (_("Error inserting file %s on the OpenOffice document: %s") % (docs, exception))

    def convertByPath(self, inputFile, outputFile):
        inputUrl = self._toFileUrl(inputFile)
        outputUrl = self._toFileUrl(outputFile)
        props = self._toProperties(Hidden=True)
        document = self.desktop.loadComponentFromURL(inputUrl, "_blank", 8, props)
        try:
            document.refresh()
        except AttributeError:
            pass
        props = self._toProperties(FilterName="writer_pdf_Export")
        try:
            document.storeToURL(outputUrl, props)
        finally:
            document.close(True)

    def _toFileUrl(self, path):
        return uno.systemPathToFileUrl(abspath(path))

    def _toProperties(self, **args):
        props = []
        for key in args:
            prop = PropertyValue()
            prop.Name = key
            prop.Value = args[key]
            props.append(prop)
        return tuple(props)

    def _restart_ooo(self):
        if not self._ooo_restart_cmd:
            self.logger.warning('No LibreOffice/OpenOffice restart script configured')
            return False
        self.logger.info('Restarting LibreOffice/OpenOffice background process')
        try:
            self.logger.info('Executing restart script "%s"' % self._ooo_restart_cmd)
            retcode = subprocess.call(self._ooo_restart_cmd, shell=True)
            if retcode == 0:
                self.logger.warning('Restart successfull')
                time.sleep(4) # Let some time for LibO/OOO to be fully started
            else:
                self.logger.error('Restart script failed with return code %d' % retcode)
        except OSError as e:
            self.logger.error('Failed to execute the restart script. OS error: %s' % e)
        return True


//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter;

import static org.artofsolving.jodconverter.office.OfficeUtils.cast;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import org.artofsolving.jodconverter.document.DocumentFamily;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.OfficeException;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.style.XStyle;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.XLineNumberingProperties;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XRefreshable;

public class StandardConversionTask extends AbstractConversionTask {

    private final DocumentFormat outputFormat;

    private Map<String,?> defaultLoadProperties;
    private DocumentFormat inputFormat;

    public StandardConversionTask(File inputFile, File outputFile, DocumentFormat outputFormat) {
        super(inputFile, outputFile);
        this.outputFormat = outputFormat;
    }

    public void setDefaultLoadProperties(Map<String, ?> defaultLoadProperties) {
        this.defaultLoadProperties = defaultLoadProperties;
    }

    public void setInputFormat(DocumentFormat inputFormat) {
        this.inputFormat = inputFormat;
    }

    @Override
    protected void modifyDocument(XComponent document, Boolean enableLineNumbering) throws OfficeException {
    	
    	if (enableLineNumbering) {
	    	XLineNumberingProperties oLNP =  UnoRuntime.queryInterface(XLineNumberingProperties.class, document);
			XPropertySet lineNumProps = oLNP.getLineNumberingProperties();
			setLineNumbering(lineNumProps);
    	}	
		
    	// Always turn off headers & footers
		try { 	
			XStyleFamiliesSupplier xSupplier = (XStyleFamiliesSupplier) UnoRuntime.queryInterface(XStyleFamiliesSupplier.class, document);
			XNameAccess xFamilies = (XNameAccess) UnoRuntime.queryInterface (XNameAccess.class, xSupplier.getStyleFamilies());
			XNameContainer xFamily = (XNameContainer) UnoRuntime.queryInterface( 
			XNameContainer.class, xFamilies.getByName("PageStyles"));
			// The style name may be "Default Style" or just "Default"
			XStyle xStyle = null;
			try {
				 xStyle = (XStyle) UnoRuntime.queryInterface(XStyle.class, xFamily.getByName("Default Style"));
			} catch (NoSuchElementException e){
				 xStyle = (XStyle) UnoRuntime.queryInterface(XStyle.class, xFamily.getByName("Default"));
			}
			XPropertySet xStyleProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xStyle);
			xStyleProps.setPropertyValue ("HeaderIsOn", Boolean.FALSE);
			xStyleProps.setPropertyValue ("FooterIsOn", Boolean.FALSE);
		} catch (WrappedTargetException|IllegalArgumentException|PropertyVetoException|UnknownPropertyException | NoSuchElementException e ) {
			e.printStackTrace();
		}
    	
        XRefreshable refreshable = cast(XRefreshable.class, document);
        if (refreshable != null) {
            refreshable.refresh();
        }
        
    }
    
    private void setLineNumbering(XPropertySet lineNumProps) {
  		try {
  			lineNumProps.setPropertyValue("IsOn", Boolean.TRUE);
  			lineNumProps.setPropertyValue("CountEmptyLines", Boolean.FALSE);
  			lineNumProps.setPropertyValue("RestartAtEachPage", Boolean.FALSE);
  			short interval = 1;
  			lineNumProps.setPropertyValue("Interval", interval);
  		} catch (WrappedTargetException|IllegalArgumentException|PropertyVetoException|UnknownPropertyException e ) {
  			e.printStackTrace();
  		}
  		
      }



    @Override
    protected Map<String,?> getLoadProperties(File inputFile) {
        Map<String,Object> loadProperties = new HashMap<String,Object>();
        if (defaultLoadProperties != null) {
            loadProperties.putAll(defaultLoadProperties);
        }
        if (inputFormat != null && inputFormat.getLoadProperties() != null) {
            loadProperties.putAll(inputFormat.getLoadProperties());
        }
        return loadProperties;
    }

    @Override
    protected Map<String,?> getStoreProperties(File outputFile, XComponent document) {
        DocumentFamily family = OfficeDocumentUtils.getDocumentFamily(document);
        return outputFormat.getStoreProperties(family);
    }


}

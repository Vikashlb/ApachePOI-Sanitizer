package com.sanitizer.DocxSanitize;

import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xslf.usermodel.*;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

public class PptxSanitize {
    public static void main(String[] args) {
        try {
            File file = new File("TestSanitize.pptx");
            OPCPackage pkg = OPCPackage.open(file);
            XMLSlideShow ppt = new XMLSlideShow(pkg);

            //To access the document metadata
            POIXMLProperties props = ppt.getProperties();
            POIXMLProperties.CoreProperties core = props.getCoreProperties();
            core.setCreator(null);               // Clears Author
            core.setTitle(null);                 // Clears Title
            core.setSubjectProperty(null);       // Clears Subject
            core.setDescription(null);           // Clears Comments/Description
            core.setKeywords(null);              // Clears Keywords
            core.setLastModifiedByUser(null);    // Clears 'Last saved by'
            core.setCategory(null);              // Clears Category

            //Remove all extended document properties
            POIXMLProperties.ExtendedProperties ext = props.getExtendedProperties();
            ext.setCompany(null);                // Clears Company
            ext.setManager(null);                // Clears Manager
            ext.setHyperlinkBase(null);          // Clears HyperLinkBase

            //Remove all custom document properties
            POIXMLProperties.CustomProperties custom = props.getCustomProperties();
            if(custom!= null) {
                List<CTProperty> list = custom.getUnderlyingProperties().getPropertyList();
                list.clear();
            }

            //Read content from slides!
            int slideNumber = 1;
            for (XSLFSlide slide : ppt.getSlides()) {
                System.out.println("Slide Number :- " + slideNumber);
                for (XSLFShape shape : slide.getShapes()) {
                    if(shape instanceof XSLFTextShape textShape) {
                        String text = textShape.getText();
                        if(text !=null && !text.isEmpty()) {
                            System.out.println(" - " + text);
                        }
                    }
                }
                slideNumber++;
                System.out.println();
            }


            //Remove a slide
            ppt.removeSlide(3); // 0-based Indexing

            //Re-Order slides
            List<XSLFSlide> slides = ppt.getSlides();
            XSLFSlide thirdSlide = slides.get(2);
            ppt.setSlideOrder(thirdSlide, 4); //move the slide at index 2 to index 4

            FileOutputStream fos = new FileOutputStream("Sanitized.pptx");
            ppt.write(fos);
            fos.close();
            ppt.close();
            pkg.close();
        }
        catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
}

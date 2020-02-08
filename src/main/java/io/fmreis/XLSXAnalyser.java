package io.fmreis;

/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.extractor.XSSFEventBasedExcelExtractor;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;

/**
 * A rudimentary XLSX -> CSV processor modeled on the
 * POI sample program XLS2CSVmra from the package
 * org.apache.poi.hssf.eventusermodel.examples.
 * As with the HSSF version, this tries to spot missing
 *  rows and cells, and output empty entries for them.
 * <p>
 * Data sheets are read using a SAX parser to keep the
 * memory footprint relatively small, so this should be
 * able to read enormous workbooks.  The styles table and
 * the shared-string table must be kept in memory.  The
 * standard POI styles table class is used, but a custom
 * (read-only) class is used for the shared string table
 * because the standard POI SharedStringsTable grows very
 * quickly with the number of unique strings.
 * <p>
 * For a more advanced implementation of SAX event parsing
 * of XLSX files, see {@link XSSFEventBasedExcelExtractor}
 * and {@link XSSFSheetXMLHandler}. Note that for many cases,
 * it may be possible to simply use those with a custom
 * {@link SheetContentsHandler} and no SAX code needed of
 * your own!
 */
public class XLSXAnalyser {

    public class MySAXTerminatorException extends RuntimeException {

    }

    private class SheetAnalyzer implements SheetContentsHandler {

        @Override
        public void startRow(int rowNum) {
            //reset the length for the most wide row
            currentRowColumnCount = 0;
        }

        @Override
        public void endRow(int rowNum) {
            //set new value if new value is greater than the current most wide row
            if(currentRowColumnCount > globalMaxColumnCount){
                globalMaxColumnCount = currentRowColumnCount;
            }
            throw new MySAXTerminatorException();
        }

        @Override
        public void cell(String cellReference, String formattedValue,
                         XSSFComment comment) {
            currentRowColumnCount++;

        }
    }


    ///////////////////////////////////////

    private final OPCPackage xlsxPackage;
    private int globalMaxColumnCount = 0;
    private int currentRowColumnCount = 0;


    /**
     * Creates a new XLSXAnalyser
     *
     * @param pkg        The XLSX package to process
     */
    public XLSXAnalyser(OPCPackage pkg) throws OpenXML4JException, SAXException, IOException {
        this.xlsxPackage = pkg;
        this.process();
    }

    @SuppressWarnings("Duplicates")
    public void processSheet(
            Styles styles,
            SharedStrings strings,
            SheetContentsHandler sheetHandler,
            InputStream sheetInputStream) throws IOException, SAXException {

                DataFormatter formatter = new DataFormatter();
                InputSource sheetSource = new InputSource(sheetInputStream);
                try {
                    XMLReader sheetParser = SAXHelper.newXMLReader();
                    ContentHandler handler = new XSSFSheetXMLHandler(
                            styles, null, strings, sheetHandler, formatter, false);
                    sheetParser.setContentHandler(handler);
                    sheetParser.parse(sheetSource);
                } catch(ParserConfigurationException e) {
                    throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
                }
    }

    @SuppressWarnings("Duplicates")
    public void process() throws IOException, OpenXML4JException, SAXException {
        long inicio = System.currentTimeMillis();

        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int index = 0;

        while (iter.hasNext()) {
            try (InputStream stream = iter.next()) {
                String sheetName = iter.getSheetName();
                try {
                    processSheet(styles, strings, new SheetAnalyzer(), stream);
                } catch (MySAXTerminatorException e){
                    System.out.println("Sheet " + sheetName + ": " + currentRowColumnCount);
                    System.out.println("Global : " + globalMaxColumnCount);
                }
            }
            ++index;
        }
        System.out.println((System.currentTimeMillis() - inicio) / 1000 + " segundos for analysis");
    }

    public int getMinimumCols() {
        return globalMaxColumnCount;
    }

    public static int getMinimumCols(String arg) throws IOException, OpenXML4JException, SAXException {
        File xlsxFile = new File(arg);

        // The package open is instantaneous, as it should be.
        try (OPCPackage opcPackage = OPCPackage.open(xlsxFile.getPath(), PackageAccess.READ)) {
            XLSXAnalyser xlsxAnalyser = new XLSXAnalyser(opcPackage);
            return xlsxAnalyser.globalMaxColumnCount;
        }
    }


    @SuppressWarnings("Duplicates")
    public static void main(String[] args) throws Exception {
        System.out.println(getMinimumCols("/home/fmreis/IdeaProjects/xlsx2csv/src/main/resources/big.xlsx"));
    }


}


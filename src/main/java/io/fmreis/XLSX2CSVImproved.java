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
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
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
import java.sql.Date;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Locale;

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
public class XLSX2CSVImproved {
    /**
     * Uses the XSSF Event SAX helpers to do most of the work
     *  of parsing the Sheet XML, and outputs the contents
     *  as a (basic) CSV.
     */
    @SuppressWarnings("Duplicates")
    private class SheetToCSV implements SheetContentsHandler {
        private boolean firstCellOfRow;
        private int currentRow = -1;
        private int currentCol = startColumn;

        private void outputMissingRows(int rowNum, int number) {
            if(rowNum <= 3)
                return;
            for (int i=0; i<number; i++) {
                for (int j=0; j<minColumns-1; j++) {
                    output.append(separator);
                }
                output.append('\n');
            }
        }

        @Override
        public void startRow(int rowNum) {
            // If there were gaps, output the missing rows
            outputMissingRows(rowNum, rowNum-currentRow-1);
            // Prepare for this row
            firstCellOfRow = true;
            currentRow = rowNum;
            currentCol = startColumn;
        }

        @Override
        @SuppressWarnings("Duplicates")
        public void endRow(int rowNum) {
            // Ensure the minimum number of columns
            if(rowNum <= 2)
                return;
            for (int i=currentCol; i<minColumns; i++) {
                output.append(separator);
            }
            output.append('\n');
        }

        @Override
        @SuppressWarnings("Duplicates")
        public void cell(String cellReference, String formattedValue,
                         XSSFComment comment) {

            if (currentRow < 3){
                return;
            }

            if (firstCellOfRow) {
                firstCellOfRow = false;
            } else {
                output.append(separator);
            }

            // gracefully handle missing CellRef here in a similar way as XSSFCell does
            if(cellReference == null) {
                cellReference = new CellAddress(currentRow, currentCol).formatAsString();
            }

            // Did we miss any cells?
            int thisCol = (new CellReference(cellReference)).getCol();
            int missedCols = thisCol - currentCol - 1;
            for (int i=0; i<missedCols; i++) {
                output.append(separator);
            }
            currentCol = thisCol;

            // Number or string?
            try {
                //noinspection ResultOfMethodCallIgnored
                //Assumes is US
                Double number = NumberFormat.getInstance(Locale.US).parse(formattedValue).doubleValue();
                output.append(numberFormat.format(number));

            } catch (ParseException e1) {

                String pattern = "yyyy-mm-dd";
                DateFormat dateFormat = new SimpleDateFormat(pattern);

                try {
                    Date parsedDate = (Date) dateFormat.parse(formattedValue);
                    System.out.println("The parsed date is: " + parsedDate);
                } catch (ParseException e2) {
                    output.append(formattedValue);
                }
            }
        }
    }


    ///////////////////////////////////////

    private final OPCPackage xlsxPackage;

    /**
     * Number of columns to read starting with leftmost
     */
    private final int minColumns;

    /**
     * Destination for data
     */
    private final PrintStream output;


    private final char separator;

    private final int startColumn = 0;

    private final String lang;

    private NumberFormat numberFormat;

    /**
     * Creates a new XLSX -> CSV examples
     *
     * @param pkg        The XLSX package to process
     * @param output     The PrintStream to output the CSV to
     * @param minColumns The minimum number of columns to output, or -1 for no minimum
     */
    public XLSX2CSVImproved(OPCPackage pkg, PrintStream output, int minColumns, char separator, String lang) {
        this.xlsxPackage = pkg;
        this.output = output;
        this.minColumns = minColumns;
        this.separator = separator;
        this.lang = lang;
    }

    /**
     * Parses and shows the content of one sheet
     * using the specified styles and shared-strings tables.
     *
     * @param styles The table of styles that may be referenced by cells in the sheet
     * @param strings The table of strings that may be referenced by cells in the sheet
     * @param sheetInputStream The stream to read the sheet-data from.

     * @exception IOException An IO exception from the parser,
     *            possibly from a byte stream or character stream
     *            supplied by the application.
     * @throws SAXException if parsing the XML data fails.
     */
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

    /**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     */
    @SuppressWarnings("Duplicates")
    public void process() throws IOException, OpenXML4JException, SAXException {
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int index = 0;

        this.numberFormat = "EN".equals(this.lang) ? NumberFormat.getInstance(Locale.US) : NumberFormat.getInstance(Locale.FRANCE);

        while (iter.hasNext()) {
            try (InputStream stream = iter.next()) {
                String sheetName = iter.getSheetName();
                processSheet(styles, strings, new SheetToCSV(), stream);
            }
            ++index;
        }
    }

    public static void main(String[] args) throws Exception {

        File xlsxFile = new File("/home/fmreis/IdeaProjects/xlsx2csv/src/main/resources/poi_test_columns.xlsx");

        char separator = ';';
        String lang = "EN";

        // The package open is instantaneous, as it should be.
        try (OPCPackage opcPackage = OPCPackage.open(xlsxFile.getPath(), PackageAccess.READ)) {
            XLSXAnalyser xlsxAnalyser = new XLSXAnalyser(opcPackage);
            XLSX2CSVImproved xlsx2csv = new XLSX2CSVImproved(opcPackage, System.out, xlsxAnalyser.getMinimumCols(), separator, lang);
            xlsx2csv.process();
        }
    }
}


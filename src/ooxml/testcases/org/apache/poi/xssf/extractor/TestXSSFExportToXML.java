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

package org.apache.poi.xssf.extractor;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.InetAddress;
import java.net.ServerSocket;
import java.net.Socket;
import java.util.Date;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.TempFile;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.XSSFTestDataSamples;
import org.apache.poi.xssf.model.MapInfo;
import org.apache.poi.xssf.usermodel.XSSFMap;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.EntityResolver;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.SAXParseException;

import junit.framework.TestCase;

/**
 * @author Roberto Manicardi
 */
public final class TestXSSFExportToXML extends TestCase {

    public void testExportToXML() throws Exception {

		XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("CustomXMLMappings.xlsx");

        boolean found = false;
		for (POIXMLDocumentPart p : wb.getRelations()) {

			if (!(p instanceof MapInfo)) {
				continue;
			}
			MapInfo mapInfo = (MapInfo) p;

			XSSFMap map = mapInfo.getXSSFMapById(1);
			XSSFExportToXml exporter = new XSSFExportToXml(map);
			ByteArrayOutputStream os = new ByteArrayOutputStream();
			exporter.exportToXML(os, true);
			String xml = os.toString("UTF-8");

			assertNotNull(xml);
			assertFalse(xml.equals(""));

			String docente = xml.split("<DOCENTE>")[1].split("</DOCENTE>")[0].trim();
			String nome = xml.split("<NOME>")[1].split("</NOME>")[0].trim();
			String tutor = xml.split("<TUTOR>")[1].split("</TUTOR>")[0].trim();
			String cdl = xml.split("<CDL>")[1].split("</CDL>")[0].trim();
			String durata = xml.split("<DURATA>")[1].split("</DURATA>")[0].trim();
			String argomento = xml.split("<ARGOMENTO>")[1].split("</ARGOMENTO>")[0].trim();
			String progetto = xml.split("<PROGETTO>")[1].split("</PROGETTO>")[0].trim();
			String crediti = xml.split("<CREDITI>")[1].split("</CREDITI>")[0].trim();

			assertEquals("ro", docente);
			assertEquals("ro", nome);
			assertEquals("ds", tutor);
			assertEquals("gs", cdl);
			assertEquals("g", durata);
			assertEquals("gvvv", argomento);
			assertEquals("aaaa", progetto);
			assertEquals("aa", crediti);
			
			parseXML(xml);
	           
           found = true;
       }
       assertTrue(found);
	}

	public void testExportToXMLInverseOrder() throws Exception {

		XSSFWorkbook wb = XSSFTestDataSamples
				.openSampleWorkbook("CustomXmlMappings-inverse-order.xlsx");

		MapInfo mapInfo = null;

        boolean found = false;
		for (POIXMLDocumentPart p : wb.getRelations()) {

			if (!(p instanceof MapInfo)) {
				continue;
			}
			mapInfo = (MapInfo) p;

			XSSFMap map = mapInfo.getXSSFMapById(1);
			XSSFExportToXml exporter = new XSSFExportToXml(map);
			ByteArrayOutputStream os = new ByteArrayOutputStream();
			exporter.exportToXML(os, true);
			String xml = os.toString("UTF-8");

			assertNotNull(xml);
			assertFalse(xml.equals(""));

			String docente = xml.split("<DOCENTE>")[1].split("</DOCENTE>")[0].trim();
			String nome = xml.split("<NOME>")[1].split("</NOME>")[0].trim();
			String tutor = xml.split("<TUTOR>")[1].split("</TUTOR>")[0].trim();
			String cdl = xml.split("<CDL>")[1].split("</CDL>")[0].trim();
			String durata = xml.split("<DURATA>")[1].split("</DURATA>")[0].trim();
			String argomento = xml.split("<ARGOMENTO>")[1].split("</ARGOMENTO>")[0].trim();
			String progetto = xml.split("<PROGETTO>")[1].split("</PROGETTO>")[0].trim();
			String crediti = xml.split("<CREDITI>")[1].split("</CREDITI>")[0].trim();

			assertEquals("aa", nome);
			assertEquals("aaaa", docente);
			assertEquals("gvvv", tutor);
			assertEquals("g", cdl);
			assertEquals("gs", durata);
			assertEquals("ds", argomento);
			assertEquals("ro", progetto);
			assertEquals("ro", crediti);
			
			parseXML(xml);
	           
           found = true;
       }
       assertTrue(found);
	}

	public void testXPathOrdering() {

		XSSFWorkbook wb = XSSFTestDataSamples
				.openSampleWorkbook("CustomXmlMappings-inverse-order.xlsx");

		MapInfo mapInfo = null;

        boolean found = false;
		for (POIXMLDocumentPart p : wb.getRelations()) {

			if (p instanceof MapInfo) {
				mapInfo = (MapInfo) p;

				XSSFMap map = mapInfo.getXSSFMapById(1);
				XSSFExportToXml exporter = new XSSFExportToXml(map);

				assertEquals(1, exporter.compare("/CORSO/DOCENTE", "/CORSO/NOME"));
				assertEquals(-1, exporter.compare("/CORSO/NOME", "/CORSO/DOCENTE"));
			}
	           
           found = true;
       }
       assertTrue(found);
	}

	public void testMultiTable() throws Exception {

		XSSFWorkbook wb = XSSFTestDataSamples
				.openSampleWorkbook("CustomXMLMappings-complex-type.xlsx");

		boolean found = false;
		for (POIXMLDocumentPart p : wb.getRelations()) {

			if (p instanceof MapInfo) {
				MapInfo mapInfo = (MapInfo) p;

				XSSFMap map = mapInfo.getXSSFMapById(2);

				assertNotNull(map);

				XSSFExportToXml exporter = new XSSFExportToXml(map);
				ByteArrayOutputStream os = new ByteArrayOutputStream();
				exporter.exportToXML(os, true);
				String xml = os.toString("UTF-8");

				assertNotNull(xml);

				String[] regexConditions = {
						"<MapInfo", "</MapInfo>",
						"<Schema ID=\"1\" Namespace=\"\" SchemaRef=\"\"/>",
						"<Schema ID=\"4\" Namespace=\"\" SchemaRef=\"\"/>",
						"DataBinding",
						"Map Append=\"false\" AutoFit=\"false\" ID=\"1\"",
						"Map Append=\"false\" AutoFit=\"false\" ID=\"5\"",
				};

				for (String condition : regexConditions) {
					Pattern pattern = Pattern.compile(condition);
					Matcher matcher = pattern.matcher(xml);
					assertTrue(matcher.find());
				}
			}
	           
           found = true;
       }
       assertTrue(found);
	}

    public void testExportToXMLSingleAttributeNamespace() throws Exception {
        // TODO: Fails, but I don't know if it is ok or not...
        
//        XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("CustomXMLMapping-singleattributenamespace.xlsx");
//
//        for (XSSFMap map : wb.getCustomXMLMappings()) {
//            XSSFExportToXml exporter = new XSSFExportToXml(map);
//
//            ByteArrayOutputStream os = new ByteArrayOutputStream();
//            exporter.exportToXML(os, true);
//        }
    }
	
    public void test55850ComplexXmlExport() throws Exception {

        XSSFWorkbook wb = XSSFTestDataSamples
                .openSampleWorkbook("55850.xlsx");

        boolean found = false;
        for (POIXMLDocumentPart p : wb.getRelations()) {

            if (!(p instanceof MapInfo)) {
                continue;
            }
            MapInfo mapInfo = (MapInfo) p;

            XSSFMap map = mapInfo.getXSSFMapById(2);

            assertNotNull("XSSFMap is null", map);

            XSSFExportToXml exporter = new XSSFExportToXml(map);
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            exporter.exportToXML(os, true);
            String xmlData = os.toString("UTF-8");

            assertNotNull(xmlData);
            assertFalse(xmlData.equals(""));

            String a = xmlData.split("<A>")[1].split("</A>")[0].trim();
            String b = a.split("<B>")[1].split("</B>")[0].trim();
            String c = b.split("<C>")[1].split("</C>")[0].trim();
            String d = c.split("<D>")[1].split("</Dd>")[0].trim();
            String e = d.split("<E>")[1].split("</EA>")[0].trim();

            String euro = e.split("<EUR>")[1].split("</EUR>")[0].trim();
            String chf = e.split("<CHF>")[1].split("</CHF>")[0].trim();

            assertEquals("15", euro);
            assertEquals("19", chf);
            
            parseXML(xmlData);
            
            found = true;
        }
        assertTrue(found);
    }

   public void testFormulaCells_Bugzilla_55927() throws Exception {
       XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("55927.xlsx");
       
       boolean found = false;
       for (POIXMLDocumentPart p : wb.getRelations()) {
           
           if (!(p instanceof MapInfo)) {
               continue;
           }
           MapInfo mapInfo = (MapInfo) p;
           
           XSSFMap map = mapInfo.getXSSFMapById(1);
           
           assertNotNull("XSSFMap is null", map);
           
           XSSFExportToXml exporter = new XSSFExportToXml(map);
           ByteArrayOutputStream os = new ByteArrayOutputStream();
           exporter.exportToXML(os, true);
           String xmlData = os.toString("UTF-8");
           
           assertNotNull(xmlData);
           assertFalse(xmlData.equals(""));
           
           assertEquals("2012-01-13", xmlData.split("<DATE>")[1].split("</DATE>")[0].trim());
           assertEquals("2012-02-16", xmlData.split("<FORMULA_DATE>")[1].split("</FORMULA_DATE>")[0].trim());
           
           parseXML(xmlData);
           
           found = true;
       }
       assertTrue(found);
   }

   public void testFormulaCells_Bugzilla_55926() throws Exception {
       XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("55926.xlsx");

       boolean found = false;
       for (POIXMLDocumentPart p : wb.getRelations()) {

           if (!(p instanceof MapInfo)) {
               continue;
           }
           MapInfo mapInfo = (MapInfo) p;

           XSSFMap map = mapInfo.getXSSFMapById(1);

           assertNotNull("XSSFMap is null", map);

           XSSFExportToXml exporter = new XSSFExportToXml(map);
           ByteArrayOutputStream os = new ByteArrayOutputStream();
           exporter.exportToXML(os, true);
           String xmlData = os.toString("UTF-8");

           assertNotNull(xmlData);
           assertFalse(xmlData.equals(""));
           
           String a = xmlData.split("<A>")[1].split("</A>")[0].trim();
           String doubleValue = a.split("<DOUBLE>")[1].split("</DOUBLE>")[0].trim();
           String stringValue = a.split("<STRING>")[1].split("</STRING>")[0].trim();
           
           assertEquals("Hello World", stringValue);
           assertEquals("5.1", doubleValue);
           
           parseXML(xmlData);
           
           found = true;
       }
       assertTrue(found);
   }
   
   @Test
   public void testXmlExportIgnoresEmptyCells_Bugzilla_55924() throws Exception {

       XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("55924.xlsx");

       boolean found = false;
       for (POIXMLDocumentPart p : wb.getRelations()) {

           if (!(p instanceof MapInfo)) {
               continue;
           }
           MapInfo mapInfo = (MapInfo) p;

           XSSFMap map = mapInfo.getXSSFMapById(1);

           assertNotNull("XSSFMap is null", map);

           XSSFExportToXml exporter = new XSSFExportToXml(map);
           ByteArrayOutputStream os = new ByteArrayOutputStream();
           exporter.exportToXML(os, true);
           String xmlData = os.toString("UTF-8");

           assertNotNull(xmlData);
           assertFalse(xmlData.equals(""));

           String a = xmlData.split("<A>")[1].split("</A>")[0].trim();
           String euro = a.split("<EUR>")[1].split("</EUR>")[0].trim();
           assertEquals("1",euro);
           
           parseXML(xmlData);
           
           found = true;
       }
       assertTrue(found);
   }

   public void testXmlExportSchemaWithXSAllTag_Bugzilla_56169() throws Exception {
       XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("56169.xlsx");

       for (XSSFMap map : wb.getCustomXMLMappings()) {
           XSSFExportToXml exporter = new XSSFExportToXml(map);

           ByteArrayOutputStream os = new ByteArrayOutputStream();
           exporter.exportToXML(os, true);
           String xmlData = os.toString("UTF-8");

           assertNotNull(xmlData);
           assertTrue(!xmlData.equals(""));

           String a = xmlData.split("<A>")[1].split("</A>")[0].trim();
           String a_b = a.split("<B>")[1].split("</B>")[0].trim();
           String a_b_c = a_b.split("<C>")[1].split("</C>")[0].trim();
           String a_b_c_e = a_b_c.split("<E>")[1].split("</EA>")[0].trim();
           String a_b_c_e_euro = a_b_c_e.split("<EUR>")[1].split("</EUR>")[0].trim();
           String a_b_c_e_chf = a_b_c_e.split("<CHF>")[1].split("</CHF>")[0].trim();

           assertEquals("1", a_b_c_e_euro);
           assertEquals("2", a_b_c_e_chf);

           String a_b_d = a_b.split("<D>")[1].split("</Dd>")[0].trim();
           String a_b_d_e = a_b_d.split("<E>")[1].split("</EA>")[0].trim();

           String a_b_d_e_euro = a_b_d_e.split("<EUR>")[1].split("</EUR>")[0].trim();
           String a_b_d_e_chf = a_b_d_e.split("<CHF>")[1].split("</CHF>")[0].trim();

           assertEquals("3", a_b_d_e_euro);
           assertEquals("4", a_b_d_e_chf);
       }
   }
   
   public void testXmlExportCompare_Bug_55923() throws Exception {
       XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("55923.xlsx");

       boolean found = false;
       for (POIXMLDocumentPart p : wb.getRelations()) {

           if (!(p instanceof MapInfo)) {
               continue;
           }
           MapInfo mapInfo = (MapInfo) p;

           XSSFMap map = mapInfo.getXSSFMapById(4);

           assertNotNull("XSSFMap is null", map);

           XSSFExportToXml exporter = new XSSFExportToXml(map);
           assertEquals(0, exporter.compare("", ""));
           assertEquals(0, exporter.compare("/", "/"));
           assertEquals(0, exporter.compare("//", "//"));
           assertEquals(0, exporter.compare("/a/", "/b/"));
           
           assertEquals(-1, exporter.compare("/ns1:Entry/ns1:A/ns1:B/ns1:C/ns1:E/ns1:EUR",
                                           "/ns1:Entry/ns1:A/ns1:B/ns1:C/ns1:E/ns1:CHF"));
           
           found = true;
       }
       assertTrue(found);
   }
   
   public void testXmlExportSchemaOrderingBug_Bugzilla_55923() throws Exception {
       XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("55923.xlsx");

       boolean found = false;
       for (POIXMLDocumentPart p : wb.getRelations()) {

           if (!(p instanceof MapInfo)) {
               continue;
           }
           MapInfo mapInfo = (MapInfo) p;

           XSSFMap map = mapInfo.getXSSFMapById(4);

           assertNotNull("XSSFMap is null", map);

           XSSFExportToXml exporter = new XSSFExportToXml(map);
           ByteArrayOutputStream os = new ByteArrayOutputStream();
           exporter.exportToXML(os, true);
           String xmlData = os.toString("UTF-8");

           assertNotNull(xmlData);
           assertFalse(xmlData.equals(""));
           
           String a = xmlData.split("<A>")[1].split("</A>")[0].trim();
           String a_b = a.split("<B>")[1].split("</B>")[0].trim();
           String a_b_c = a_b.split("<C>")[1].split("</C>")[0].trim();
           String a_b_c_e = a_b_c.split("<E>")[1].split("</EA>")[0].trim();
           String a_b_c_e_euro = a_b_c_e.split("<EUR>")[1].split("</EUR>")[0].trim();
           String a_b_c_e_chf = a_b_c_e.split("<CHF>")[1].split("</CHF>")[0].trim();
           
           assertEquals("1",a_b_c_e_euro);
           assertEquals("2",a_b_c_e_chf);
           
           String a_b_d = a_b.split("<D>")[1].split("</Dd>")[0].trim();
           String a_b_d_e = a_b_d.split("<E>")[1].split("</EA>")[0].trim();
           
           String a_b_d_e_euro = a_b_d_e.split("<EUR>")[1].split("</EUR>")[0].trim();
           String a_b_d_e_chf = a_b_d_e.split("<CHF>")[1].split("</CHF>")[0].trim();
           
           assertEquals("3",a_b_d_e_euro);
           assertEquals("4",a_b_d_e_chf);
           
           found = true;
       }
       assertTrue(found);
   }
   
   private void parseXML(String xmlData) throws IOException, SAXException, ParserConfigurationException {
       DocumentBuilderFactory docBuilderFactory = XMLHelper.getDocumentBuilderFactory();
       docBuilderFactory.setNamespaceAware(true);
       docBuilderFactory.setValidating(false);
       DocumentBuilder docBuilder = docBuilderFactory.newDocumentBuilder();
       docBuilder.setEntityResolver(new DummyEntityResolver());

       docBuilder.parse(new ByteArrayInputStream(xmlData.getBytes("UTF-8")));
   }

   private static class DummyEntityResolver implements EntityResolver
   {
       @Override
       public InputSource resolveEntity(String publicId, String systemId) throws SAXException, IOException
       {
           return null;
       }
   }
   
   public void testExportDataTypes() throws Exception {
       XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("55923.xlsx");
       
       Sheet sheet = wb.getSheetAt(0);
       Row row = sheet.getRow(0);
       
       Cell cString = row.createCell(0);
       cString.setCellValue("somestring");
       cString.setCellType(CellType.STRING);
       
       Cell cBoolean = row.createCell(1);
       cBoolean.setCellValue(true);
       cBoolean.setCellType(CellType.BOOLEAN);
       
       Cell cError = row.createCell(2);
       cError.setCellType(CellType.ERROR);
       
       Cell cFormulaString = row.createCell(3);
       cFormulaString.setCellFormula("A1");
       cFormulaString.setCellType(CellType.FORMULA);
       
       Cell cFormulaNumeric = row.createCell(4);
       cFormulaNumeric.setCellFormula("F1");
       cFormulaNumeric.setCellType(CellType.FORMULA);
       
       Cell cNumeric = row.createCell(5);
       cNumeric.setCellValue(1.2);
       cNumeric.setCellType(CellType.NUMERIC);
       
       Cell cDate = row.createCell(6);
       cDate.setCellValue(new Date());
       cDate.setCellType(CellType.NUMERIC);
       
       boolean found = false;
       for (POIXMLDocumentPart p : wb.getRelations()) {

           if (!(p instanceof MapInfo)) {
               continue;
           }
           MapInfo mapInfo = (MapInfo) p;

           XSSFMap map = mapInfo.getXSSFMapById(4);

           assertNotNull("XSSFMap is null", map);

           XSSFExportToXml exporter = new XSSFExportToXml(map);
           ByteArrayOutputStream os = new ByteArrayOutputStream();
           exporter.exportToXML(os, true);
           String xmlData = os.toString("UTF-8");

           assertNotNull(xmlData);
           assertFalse(xmlData.equals(""));
           
           parseXML(xmlData);
           
           found = true;
       }
       assertTrue(found);
   }

   public void testValidateFalse() throws Exception {
       XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("55923.xlsx");

       boolean found = false;
       for (POIXMLDocumentPart p : wb.getRelations()) {

           if (!(p instanceof MapInfo)) {
               continue;
           }
           MapInfo mapInfo = (MapInfo) p;

           XSSFMap map = mapInfo.getXSSFMapById(4);

           assertNotNull("XSSFMap is null", map);

           XSSFExportToXml exporter = new XSSFExportToXml(map);
           ByteArrayOutputStream os = new ByteArrayOutputStream();
           exporter.exportToXML(os, false);
           String xmlData = os.toString("UTF-8");

           assertNotNull(xmlData);
           assertFalse(xmlData.equals(""));
           
           parseXML(xmlData);
           
           found = true;
       }
       assertTrue(found);
   }

   public void testRefElementsInXmlSchema_Bugzilla_56730() throws Exception {
       XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("56730.xlsx");
       
       boolean found = false;
       for (POIXMLDocumentPart p : wb.getRelations()) {
           
           if (!(p instanceof MapInfo)) {
               continue;
           }
           MapInfo mapInfo = (MapInfo) p;
           
           XSSFMap map = mapInfo.getXSSFMapById(1);
           
           assertNotNull("XSSFMap is null", map);
           
           XSSFExportToXml exporter = new XSSFExportToXml(map);
           ByteArrayOutputStream os = new ByteArrayOutputStream();
           exporter.exportToXML(os, true);
           String xmlData = os.toString("UTF-8");
           
           assertNotNull(xmlData);
           assertFalse(xmlData.equals(""));
           
           assertEquals("2014-12-31", xmlData.split("<DATE>")[1].split("</DATE>")[0].trim());
           assertEquals("12.5", xmlData.split("<REFELEMENT>")[1].split("</REFELEMENT>")[0].trim());

           parseXML(xmlData);

           found = true;
       }
       assertTrue(found);
   }

    public void testXXEInSchema() throws Exception {
        XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("xxe_in_schema.xlsx");
        try {
            // exporting a table cell of a non-mapped column fails before the schema validation
            // (see bug 61281), so drop the empty one in D10
            Row row = wb.getSheetAt(0).getRow(9);
            row.removeCell(row.getCell(3));

            for (XSSFMap map : wb.getCustomXMLMappings()) {
                XSSFExportToXml exporter = new XSSFExportToXml(map);
                ByteArrayOutputStream bos = new ByteArrayOutputStream();
                exporter.exportToXML(bos, true);
            }
            fail("SAXParseException expected, the schema redefines an external schema");
        } catch (SAXParseException e) {
            // expected, the external schema is not loaded
        } finally {
            wb.close();
        }
    }

    public void testExternalSchemaReferencesAreNotLoaded() throws Exception {
        final ServerSocket server = new ServerSocket(0, 50, InetAddress.getByName("127.0.0.1"));
        final AtomicInteger requests = new AtomicInteger();
        Thread listener = new Thread("XXE schema listener") {
            @Override
            public void run() {
                try {
                    while (true) {
                        Socket socket = server.accept();
                        requests.incrementAndGet();
                        socket.close();
                    }
                } catch (IOException e) {
                    // server socket closed
                }
            }
        };
        listener.setDaemon(true);
        listener.start();
        try {
            for (String reference : new String[]{"redefine", "include", "import"}) {
                String location = "http://127.0.0.1:" + server.getLocalPort() + "/" + reference + ".xsd";
                XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("CustomXMLMappings.xlsx");
                try {
                    XSSFMap map = wb.getCustomXMLMappings().iterator().next();
                    addSchemaReference(map, reference, location);

                    SAXParseException failure = null;
                    try {
                        new XSSFExportToXml(map).exportToXML(new ByteArrayOutputStream(), true);
                    } catch (SAXParseException e) {
                        failure = e;
                    }
                    assertEquals("xs:" + reference + " must not load " + location, 0, requests.get());
                    assertNotNull("validation must fail for a schema with an external xs:" + reference, failure);
                } finally {
                    wb.close();
                }
            }
        } finally {
            server.close();
            listener.join(10000);
        }
    }

    public void testLocalSchemaFileIsNotIncluded() throws Exception {
        File xsd = TempFile.createTempFile("xxe_in_schema", ".xsd");
        try {
            OutputStream out = new FileOutputStream(xsd);
            try {
                out.write(("<xsd:schema xmlns:xsd=\"" + XMLConstants.W3C_XML_SCHEMA_NS_URI + "\">"
                        + "<xsd:simpleType name=\"LocalFileType\"><xsd:restriction base=\"xsd:string\"/></xsd:simpleType>"
                        + "</xsd:schema>").getBytes("UTF-8"));
            } finally {
                out.close();
            }

            XSSFWorkbook wb = XSSFTestDataSamples.openSampleWorkbook("CustomXMLMappings.xlsx");
            try {
                XSSFMap map = wb.getCustomXMLMappings().iterator().next();
                addSchemaReference(map, "include", xsd.toURI().toString());

                // the embedded schema now depends on the local file, so it can only be valid if the file is read
                NodeList elements = ((Element) map.getSchema()).getElementsByTagNameNS(XMLConstants.W3C_XML_SCHEMA_NS_URI, "element");
                boolean found = false;
                for (int i = 0; i < elements.getLength(); i++) {
                    Element element = (Element) elements.item(i);
                    if ("NOME".equals(element.getAttribute("name"))) {
                        element.setAttribute("type", "LocalFileType");
                        found = true;
                    }
                }
                assertTrue(found);

                try {
                    new XSSFExportToXml(map).exportToXML(new ByteArrayOutputStream(), true);
                    fail("the local schema file " + xsd + " must not be included");
                } catch (SAXParseException e) {
                    // expected, the local schema file is not loaded
                }
            } finally {
                wb.close();
            }
        } finally {
            xsd.delete();
        }
    }

    private static void addSchemaReference(XSSFMap map, String reference, String location) {
        Element schema = (Element) map.getSchema();
        String prefix = schema.getPrefix();
        Element ref = schema.getOwnerDocument().createElementNS(XMLConstants.W3C_XML_SCHEMA_NS_URI,
                (prefix == null ? "" : prefix + ":") + reference);
        if ("import".equals(reference)) {
            ref.setAttribute("namespace", "urn:poi:xxe");
        }
        ref.setAttribute("schemaLocation", location);
        schema.insertBefore(ref, schema.getFirstChild());
    }
}

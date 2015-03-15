
import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.URL;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.print.Doc;
import javax.print.DocFlavor;
import javax.print.DocPrintJob;
import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.SimpleDoc;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.HashPrintServiceAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.PrintServiceAttributeSet;
import javax.print.attribute.standard.Copies;
import javax.print.attribute.standard.MediaSizeName;
import javax.print.attribute.standard.PrinterName;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JRExporterParameter;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperPrintManager;
import net.sf.jasperreports.engine.JasperRunManager;
import net.sf.jasperreports.engine.export.JRPrintServiceExporter;
import net.sf.jasperreports.engine.export.JRPrintServiceExporterParameter;
import net.sf.jasperreports.engine.query.JRXPathQueryExecuterFactory;
import net.sf.jasperreports.engine.util.JRXmlUtils;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
/**
 * @author Teodor Danciu (teodord@users.sourceforge.net)
 * @version $Id: XmlDataSourceApp.java,v 1.13 2006/04/19 10:26:14 teodord Exp $
 */
public class BuildReport7
{

	/**
	 * Re-Write based on BuildReport wrapper. This extends the framework to run
	 * JRXML's with parameters and SQL query using Evolution Datasources.
	 *
	 */
	private static final String TASK_PRINT = "print";
	private static final String TASK_EMAIL = "email";
	private static final String TASK_PDF_P = "pdfprint";
	private static final String TASK_PDF_E = "pdfemail";
	private static final String TASK_PDF   = "pdf";

	// Program configuration settings go in here
	private static Properties configProps = new Properties();

	private static void replaceTxt(String oldstring, String newstring, File in, File out) throws Exception {
		try {
			BufferedReader reader = new BufferedReader(new FileReader(in));
			PrintWriter writer = new PrintWriter(new FileWriter(out));
			String line = null;
			while ((line = reader.readLine()) != null) {
				//TODO Change to replaceALL(); with a regex
				writer.println(line.replace(oldstring,newstring));
			}
			reader.close();
			writer.close();
		}
		catch( IOException e )
		{
			System.out.println("replaceTxt - IOException: " + e.getMessage()  );
			throw e;
		}
	}

	/**
	 *
	 */
	public static void main(String[] args)
	{
		// Setup Log4J Logging
		//org.apache.log4j.BasicConfigurator.configure();
	    
		// Read config options from properties file
		try {
			String configFilename = System.getenv("CONFIG_FILENAME");
			if( configFilename == null) {
				usage();
				System.err.println("ERROR: Environment variable CONFIG_FILENAME not set");
				
				return;
			}
			
			InputStream inStream = new FileInputStream(configFilename);

			configProps.load(inStream);
		    if (inStream != null)
		        inStream.close();

		} catch (IOException e) {
			System.err.println("ERROR: " + e.getMessage());
			return;
		} 
				
		String taskName = null;				// -T What to perfrom
		String fileName = null;				// -F Input report format .jasper
		String dataName = null;				// -S Input XML file
		String outputName = null;			// -O What to do with the output eg c:\file.pdf
		String printerName = null;			// -P Printer queue
		int printCopies = new Integer(0);	// -C Print Copies
		String emailName = null;			// -ET Email to address (could be XML to feed from source file)
		String emailFrom = null;			// -EF Email From address
		String[] parameters = null;			// -V Parameters as string eg. Param1=Val1
		
		if(args.length == 0)
		{
		        usage();
				System.err.println("ERROR: No command line arguments supplied");
		        return;
		}
		
		int k = 0;
		while ( args.length > k )
		{
		        if ( args[k].startsWith("-T") )
		                taskName = args[k].substring(2);
		        if ( args[k].startsWith("-F") )
		                fileName = args[k].substring(2);
		        if ( args[k].startsWith("-S") )
		                dataName = args[k].substring(2);
		        if ( args[k].startsWith("-O") )
		                outputName = args[k].substring(2);
		        if ( args[k].startsWith("-P") )
		                printerName = args[k].substring(2);
		        if ( args[k].startsWith("-C") )
		                printCopies = Integer.parseInt( args[k].substring(2) );
		        if ( args[k].startsWith("-ET") )
		                emailName = args[k].substring(3);
		        if ( args[k].startsWith("-EF") )
		                emailFrom = args[k].substring(3);
		        if ( args[k].startsWith("-V") )
				{
					String s = args[k].substring(2);
					parameters = s.split(":");
				}
		        k++;
		}
		
		
		try
		{
		
			// Only do this cleanup for XML input files
			if( parameters == null ) {
				File originalFile = new File( dataName );
				File stagingFile = new File( dataName + ".cleaned.xml" );
				replaceTxt("&","&amp;", originalFile , stagingFile );
				originalFile.delete();
				stagingFile.renameTo( originalFile );
			}
			
			long start = System.currentTimeMillis();
		
			if (TASK_PRINT.equals(taskName))
			{
				sendToPrinter(printerName, dataName, fileName, printCopies, parameters);
				System.out.println("Total processing time : " + (System.currentTimeMillis() - start));
			
				System.exit(0);
			}
			else if (TASK_EMAIL.equals(taskName))
			{
				// TODO: Test Input params before we begin
		
				makePDF(dataName, fileName, outputName, true, parameters);
				System.out.println("PDF running time : " + (System.currentTimeMillis() - start));
		
				emailFrom = getEmailAddress( emailFrom );
		
				sendemail( emailFrom, emailName, outputName ,  JRXmlUtils.parse(new File( dataName )) , fileName );
				System.out.println("Printing time : " + (System.currentTimeMillis() - start));
				System.exit(0);
			}
			else if (TASK_PDF_P.equals(taskName))
			{ // Function to store PDF and deliver to a print queue
			
				// TODO: Test Input params before we begin		
				makePDF(dataName, fileName, outputName, true, parameters);
				System.out.println("PDF creation time : " + (System.currentTimeMillis() - start));
		
				sendToPrinter(printerName, dataName, fileName, printCopies, parameters);

				// Check if Batchcard, if so, pickup Cutsheet and print also
				if( "ubfbc.jasper".equals( fileName ))
				{
					//TODO Remove Hardcoded XML Tag name
					String suppliment = getSuppliment(dataName,"/UBF993/DOCHEADER/SALESORDER");
					
					//TODO Remove Hardcoded path, should be parameterised
					suppliment = "C:\\mtmsroot\\mtms\\DOCUMENTS\\UB\\OUTPUT\\CUTSHEET\\" + suppliment + ".pdf";

					System.out.println("Looking for cutsheet : " + suppliment);
					printFile( suppliment, printerName );
				}

				System.out.println("Total processing time : " + (System.currentTimeMillis() - start));
		
				System.exit(0);
			}
			else if (TASK_PDF_E.equals(taskName))
			{ // Function to store PDF and deliver an email document
		
				// TODO: Test Input params before we begin
		
				System.out.println("INFO: Email To: " + emailName );
				System.out.println("INFO: Email From: " + emailFrom );
		
				makePDF(dataName, fileName, outputName, true, parameters);
				System.out.println("INFO: PDF render time : " + (System.currentTimeMillis() - start));
		
				Document multiDoc = JRXmlUtils.parse(new File( dataName ));
				MultiDoc m = isMultiDoc( multiDoc, fileName );
				
				emailFrom = getEmailAddress( emailFrom );
		
				if( m.getDocCount() > 1)
				{
					//-------------------------------------------------------------------
					// Multi Document (Batch Production)
					
					System.out.println("INFO: Multipage Document detected!");
		
					// Establish this flag to force lookup of email address for XML source file
					Boolean isXMLSource = new Boolean( false );
					
					NodeList list = multiDoc.getElementsByTagName( m.getDocElement() );
		
					// Iterator for each document
					for(int x = 0 ; x < m.getDocCount() ; x++)
					{
						
						// First Job - Extract single document from multi-doc source
						Document singledoc = extractDoc( (Element)list.item( x ) , m.getDocRoot() );
		
						// Second Job - Use the document key to pull out the reference
						// for this document eg. 26014667
						NodeList singleDocNodeList = singledoc.getElementsByTagName( m.getDocKey() );
						Node subnode = singleDocNodeList.item(0);
						String text = subnode.getTextContent();
						if( text.equals("") ) {
							continue;
						}
		
						// Extract output path and re-construct file from the node in the source file
						File outputFile = new File(outputName);  
						String ofilePath= outputFile.getParent();								
						outputName =  ofilePath + "\\" + text + ".pdf";
		
						// Extract email address for this iteration
						if (emailName.length() == 3  ||  isXMLSource )
						{
		
							// Email address is either in the XML source file, or we have a requestor
							if (emailName.equals("XML") ||  isXMLSource )
							{
								
								// Reset this flag so future iterations still lookup email address for XML source file
								isXMLSource = new Boolean( true );
		
								NodeList eMailNodeList = singledoc.getElementsByTagName( "EMAIL" );
								Node eMailNode = eMailNodeList.item(0);
		
								// Protect if no EMAIL Tag found.
								if(eMailNode == null) {
									// null handling
									emailName = "";
									System.out.println("WARNING: EMail Tag was missing.");
								}
								else {
								   emailName = eMailNode.getTextContent();
								}
		
								if( emailName == "" ) {
									// No Address in file, return to sender
									emailName = emailFrom;
									System.out.println("XML Address was Blank. Using Sender address.");
								}
		
							
							} else {
		
								System.out.println("Email paramter is Evolution Requestor!");
								emailName = getEmailAddress( emailName );
							}
						}
		
						sendemail( emailFrom, emailName, outputName ,multiDoc , fileName);
					}
				
				} else {
					// Single page document code goes here
					System.out.println("INFO: Emailing Single page document!");
		
					// Check email address is a requestor eg. SYS or NKO
					if (emailName.length() == 3 )
					{
		
						// Email address is either in the XML source file, or we have a requestor
						if (emailName.equals("XML") )
						{
							System.out.println("INFO: Email paramter is in XML! Looking for <EMAIL>...</EMAIL>");
		
							// Use the node from earlier to extract single email address
							NodeList eMailNodeList = multiDoc.getElementsByTagName( "EMAIL" );
							Node eMailNode = eMailNodeList.item(0);
		
							// Protect if no EMAIL Tag found.
							if(eMailNode == null) {
								// null handling
								emailName = "";
								System.out.println("WARNING: Email Tag was missing.");
							}
							else {
							   emailName = eMailNode.getTextContent();
							}
		
							if( emailName == "" ) {
								// No Address in file, return to sender
								emailName = emailFrom;
								System.out.println("WARNING: XML Address was Blank. Using Sender address.");
							}
		
							System.out.println("INFO: Setting email address to : " + emailName);
						
						} else {
							System.out.println("INFO: Email paramter is Evolution Requestor!");
		
							emailName = getEmailAddress( emailName );
						}
					}
		
					sendemail( emailFrom, emailName, outputName , multiDoc , fileName);
				
				}
		
		
				System.out.println("Total processing time : " + (System.currentTimeMillis() - start));
				System.exit(0);
			}						
			else if (TASK_PDF.equals(taskName))
			{ // Function to store PDF and deliver an email document
		
				// TODO: Test Input params before we begin
		
				if (emailName.length() == 3 )
				{
					emailName = getEmailAddress( emailName );
				}
				
				makePDF(dataName, fileName, outputName, true, parameters);
				System.out.println("PDF running time : " + (System.currentTimeMillis() - start));
		
				System.exit(0);
			}						
			else
			{
					usage();
					System.exit(0);
			}
		}
		catch (JRException e)
		{
		        e.printStackTrace();
		        System.exit(1);
		}
		catch (Exception e)
		{
		        e.printStackTrace();
		        System.exit(1);
		}
	}

	/**
	 *  Method to retrieve a text string from an XML document
	 * @param dataName name of XML document (eg. c:\somefile.xml)
	 * @param keyName Path in XML document (eg. /DOC/BRANCH/KEY )
	 * @return
	 * @throws Exception
	 */
	private static String getSuppliment(String dataName, String keyName) throws Exception
	{
		String suppStr;
		
		try {
			// Get XML document
			Document srcDoc = JRXmlUtils.parse(new File( dataName ));
			
			// Initialize variable
			XPathExpression expr = null;
					
			// Create a XPathFactory
			XPathFactory xFactory = XPathFactory.newInstance();

			// Create a XPath object
			XPath xpath = xFactory.newXPath();

			// Compile the XPath expression
			expr = xpath.compile( keyName );  // eg. "/UBF993/DOCHEADER/SALESORDER"

			suppStr = (String) expr.evaluate(srcDoc, XPathConstants.STRING);
		}
		catch (XPathExpressionException e )
		{
			System.err.println("ERROR: Could not parse XML document. " + e.getMessage() );
			throw new Exception(e);
		}
		
		return suppStr;
	}
	
	/**
	 * Method to print an external PDF document
	 * @param fileName
	 * @param printerName
	 */
	private static void printFile(String fileName, String printerName)
	{
		FileInputStream psStream = null;

		try {
			psStream = new FileInputStream( fileName );
		} catch (FileNotFoundException ffne) {
			System.out.println("ERROR: " + fileName + " does not exist!");
			ffne.printStackTrace();
		}

		if (psStream == null) {
			return;
		}

		DocFlavor psInFormat = DocFlavor.INPUT_STREAM.AUTOSENSE;
		Doc myDoc = new SimpleDoc(psStream, psInFormat, null); 
		         
		// this step is necessary because I have several printers configured
		PrintService myPrinter = null;
		
		//list all printers and find the one selected by the user
        PrintRequestAttributeSet aset = new HashPrintRequestAttributeSet();
		PrintService[] servicesList = PrintServiceLookup.lookupPrintServices(null, aset);
		for(int l=0;l<servicesList.length;l++){
			
			if(servicesList[l].getName().trim().equals( printerName.trim() )){
				System.out.println("BINGO! Printer queue name:" + servicesList[l].getName() );
				myPrinter = servicesList[l];
				break;
			}
		}
		         
        if (myPrinter != null) {           
            DocPrintJob job = myPrinter.createPrintJob();
            try {
            job.print(myDoc, aset);
             
            } catch (Exception pe) {pe.printStackTrace();}
        } else {
            System.out.println("ERROR: Could not find printer queue " + printerName.trim() );
        }
		
	}
        /**
         * Method of parse XML feeds for multi-page documents
		 *
		 * @param  doc	Source Data to be merged as org.w3c.dom.Document object
		 * @param  key	Text string of unique document identifier eg. "UBF007"
		 *
         */
		private static Double getDocumentCount(Document doc, String key) throws Exception
		{
			
			Double docCount = new Double(0);
			try {
			
				// Initialize variable
				XPathExpression expr = null;
						
				// Create a XPathFactory
				XPathFactory xFactory = XPathFactory.newInstance();

				// Create a XPath object
				XPath xpath = xFactory.newXPath();

				// Compile the XPath expression
				expr = xpath.compile("count( " + key + " )");

				Double number = (Double) expr.evaluate(doc, XPathConstants.NUMBER);
				docCount = number;
			}
			catch (XPathExpressionException e )
			{
				System.err.println("Ops! Could not count documents: " + e.getMessage() );
				throw new Exception(e);
			}
			return docCount;
		}

        /**
         * Method to extract a single document from a multi-doc XML
		 *
		 * @param  sourceElement	Source Data to be merged as org.w3c.dom.Element object
		 * @param  key				Text string of unique document identifier eg. "UBF007"
		 *
         */
		private static Document extractDoc(Element sourceElement, String key) throws Exception
		{

			System.out.println("DEBUG(getDocumentCount): Search Key=" + key );
		
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			DocumentBuilder builder = factory.newDocumentBuilder();
			Document doc2 = builder.newDocument(); 

			try
			{
				// Create root element "UBF007"
				Element elem2 = doc2.createElement( key );
				// Append the new element at the start of the document
				doc2.appendChild(elem2);
				// perform a deep copy of 'sourceElement', adding it as child of 'elem2' from 'doc2'
				elem2.appendChild(doc2.importNode(sourceElement, true));

			}
			catch( Exception e)
			{
				System.err.println("Ops! Could not extract document: " + e.getMessage() );
				throw new Exception(e);			
			}
			
			return doc2;
		}


        /**
         * Debugging method for use with XML feeds
		 *
		 * @param  doc	Source Data to be merged as org.w3c.dom.Document object
		 *
         */
		private static void disaplyXMLDoc( Document doc ) throws Exception
		{
			TransformerFactory tranFactory = TransformerFactory.newInstance(); 
			Transformer aTransformer = tranFactory.newTransformer(); 
			Source src = new DOMSource(doc); 
			Result dest = new StreamResult(System.out); 
			aTransformer.transform(src, dest); 				

		}
			
		
        /**
         * Function to scan XML feed for multiple keys. This also sets up some key words for different documents
		 *
		 * @param  document		Source Data to be merged as org.w3c.dom.Document object
		 * @param  fileName		Template file in use (eg. Invoice.jasper)
		 *
         */
		private static MultiDoc isMultiDoc(Document document, String fileName) throws Exception
		{
			MultiDoc m = new MultiDoc();
			
			try {
					// Determine if we need to search for a multi-document
				if( fileName.endsWith("UBF_Invoice_P.Jasper") ) {
				
					m.setDocRoot( "UBF007" );
					m.setDocElement( "DESPATCH" );
					m.setDocKey( "DESPNO" );
				
					// Check Document Tags for Multi-Doc!
					m.setDocCount(  getDocumentCount( document, "/" + m.getDocRoot() + "/" + m.getDocElement()  )   );
			
				} else if( fileName.endsWith("UBF_Invoice_L.Jasper") ) {
				
					m.setDocRoot( "UBF007" );
					m.setDocElement( "DESPATCH" );
					m.setDocKey( "DESPNO" );
				
					// Check Document Tags for Multi-Doc!
					m.setDocCount(  getDocumentCount( document, "/" + m.getDocRoot() + "/" + m.getDocElement()  )   );

				} else if( fileName.endsWith("UBF_Statement.jasper") ) {
					m.setDocRoot( "DOCROOT");
					m.setDocElement( "STATEMENT");
					m.setDocKey( "CUS_NUM");

					// Check Document Tags for Multi-Doc!
					m.setDocCount( getDocumentCount( document, "/" + m.getDocRoot() + "/" + m.getDocElement() ) );
				} else if( fileName.endsWith("TKInvoices.jasper") ) {
					m.setDocRoot( "Document");
					m.setDocElement( "INVOICE");
					m.setDocKey( "InvoiceNo");

					// Check Document Tags for Multi-Doc!
					m.setDocCount( getDocumentCount( document, "/" + m.getDocRoot() + "/" + m.getDocElement() ) );
				} else if( fileName.endsWith("TKStatement.jasper") ) {
					m.setDocRoot( "Document");
					m.setDocElement( "Statement");
					m.setDocKey( "CustNo");

					// Check Document Tags for Multi-Doc!
					m.setDocCount( getDocumentCount( document, "/" + m.getDocRoot() + "/" + m.getDocElement() ) );
				} else if( fileName.endsWith("REMITTANCE.jasper") ) {
					m.setDocRoot( "EFR298");
					m.setDocElement( "Remittance");
					m.setDocKey( "RemittanceNo");

					// Check Document Tags for Multi-Doc!
					m.setDocCount( getDocumentCount( document, "/" + m.getDocRoot() + "/" + m.getDocElement() ) );
										
				} else if( fileName.endsWith("PurchaseOrder.jasper") ) {
					System.out.println("DEBUG: isMultiDoc() - Detected Purchase Order");
					
					m.setDocRoot( "docroot");
					m.setDocElement( "porder");
					m.setDocKey( "header/document");

					//TODO Currently only single PO produced. Should handle batch run PO's
					m.setDocCount( new Double(1));
					
				} else {
					m.setDocCount( new Double(1));
				}			
			} catch( Exception e) {
				System.err.println("DOH ERROR:" + e.getMessage() );				
				throw e;
			}
			return m;
		
		}
		
        /**
         * This method is the core to generate a PDF document.
		 *
		 * @param  dataname		Source Data to be merged eg. d:\mtmsroot\mtmsprt\mydata.xml
		 * @param  filename		Document Template eg. MyFormat.jasper
		 * @param  outputname 	Output file & path eg. d:\somewhere\documents\26015628.pdf
		 * @param  needlink 	
		 * @param  parameters 	String array of inputs to set the context of the report
		 * @return
         */
		private static void makePDF(String dataName, String fileName, String outputName, Boolean needLink,String[] parameters) throws JRException, Exception
		{
			
			Map<String,Object> params = new HashMap<String,Object>(); // Storage for rendering parameters.

			Document document = null;

			// Test String[] parameters to see if Empty or not. Determines if this is SQL or XML process.
			if( dataName != null && parameters == null ) {
				document = JRXmlUtils.parse(new File( dataName ));
			}
			
			MultiDoc m = isMultiDoc( document, fileName );
			
			// If Documents > 1 then split for MultiDoc
			if( m.getDocCount() > 1) 
			{
				System.out.println("Multidoc! Splitting...");
				
				NodeList list = document.getElementsByTagName( m.getDocElement() );

				// Iterator for each document
				for(int x = 0 ; x < m.getDocCount() ; x++)
				{
					// First Job - Extract single document from multi-doc source
					Document singledoc = extractDoc( (Element)list.item( x ) , m.getDocRoot() );

					// DEBUGGING ONLY disaplyXMLDoc( singledoc );

					// Second Job - Use the document key to pull out the reference
					// for this document eg. 26014667
					NodeList singleDocNodeList = singledoc.getElementsByTagName( m.getDocKey() );
					Node subnode = singleDocNodeList.item(0);
					String text = subnode.getTextContent();
					if( text.equals("") ) {
						continue;
					}

					// Extract output path and re-construct file from the node in the source file
					File outputFile = new File(outputName);  
					String ofilePath= outputFile.getParent();
					
					outputName =  ofilePath + "\\" + text + ".pdf";
					System.out.println("Producing document: " + outputName );
					
					params.put( JRXPathQueryExecuterFactory.PARAMETER_XML_DATA_DOCUMENT, singledoc);	// Use Generics
					JasperRunManager.runReportToPdfFile( fileName, outputName , params);

					makeSymLink(text,outputName );
				}

			}
			else	// Render as-is .. no more fluffing around
			{

				System.out.println("Single Document Mode");

				if( parameters != null ) {
					System.out.println("SQL datasource /w Parameters");
					System.out.println("No. of Parameters:" + parameters.length);

					//-----------------------------------------------------
					// Prepare the Parameters to run the report

					for (int ctr=0; ctr < parameters.length; ctr++) {
						// Extract the first value pair from the array
						String thisParam = parameters[ctr];

						// Split the value pair
						String[] thisParamArr = thisParam.split("=");

						// Set the value pair in a Hashmap for Jasper
						params.put(thisParamArr[0], thisParamArr[1]);
					}

					//-----------------------------------------------------
					// Prepare the Oracle connection

					Connection conn = null;
					try {
						Class.forName ("oracle.jdbc.OracleDriver");

						String url = configProps.getProperty( "jdbc.connection.url" );
						String user = configProps.getProperty( "jdbc.connection.username" );
						String pass = configProps.getProperty( "jdbc.connection.password" );

						conn = DriverManager.getConnection( url , user , pass );
					} catch(SQLException e) {
						System.err.println("ERROR:" + e.getMessage() );				
						throw e;
					}

					//-----------------------------------------------------
					// Bring it all together and create the PDF document
					
					JasperRunManager.runReportToPdfFile(fileName, outputName , params, conn);

					makeSymLink("",outputName );
				
				} else {
					System.out.println("XML Datasource");

					//-----------------------------------------------------
					// Traditional single document from XML feed

					params.put(JRXPathQueryExecuterFactory.PARAMETER_XML_DATA_DOCUMENT, document);
					JasperRunManager.runReportToPdfFile(fileName, outputName , params);

					makeSymLink("",outputName );
				}
			}	
		}

        /**
         *	Function to make a Symbolic Link. 
		 * Native methods exist in Java 1.7 however UBF still using 1.6 so using
		 * a basic command stack to do it.
		 *
		 *	@param	text		Text string of document ID
		 *	@param	outputName	Text string to the full output PDF file (source of Link)
		 *
         */
		private static void makeSymLink(String text, String outputName) throws JRException, Exception
		{

			// Flum it like it's stolen
			File outputFile = new File(outputName);  	

			// Get the source directory branch from the output file location
			String ofilePath= outputFile.getParent();

			// Trap for single-doc. Take the document number from the output filename
			if( "".equals( text ) ) {
				String fileName = outputFile.getName();
				text = fileName.substring(0, fileName.lastIndexOf('.'));
				System.out.println("SingleDoc Refernce: " + text );
			}

			// Calculate the new directory branch for the link
			String[] directories = ofilePath.substring(1).split("OUTPUT");
			String linkName = "\\" + directories[0] + "OUTPUT\\LINKS\\" + text + "." + directories[1].substring(1,4) +".pdf";

			System.out.println("Building Symlink to: " + linkName );
			
			String[] envp = {};					
			String[] cmd = {"cmd.exe" ,"/C", "mklink", linkName, outputName};
			Process p;
			p = Runtime.getRuntime().exec(cmd, envp);

			// Use this to capture any text from the console
			BufferedReader in = new BufferedReader( new InputStreamReader( p.getInputStream() )  );  

			if(p.waitFor() > 0) {
				// Something went wrong. Put it out to the screen/log
				
				String line = null;  
				while ((line = in.readLine()) != null) {  
					System.err.println("ERROR:" + line);  
				}
			} else {
				System.out.println("Symlink successful.");
			}
		}
		
        /**
         *	Function to execute report and dump down to a printer queue
		 *
		 *	@param	printerName		Text string of printer queue name
		 *	@param	dataName		Text string of full path & filename for XML source data
		 *	@param	fileName		Text string to the full output PDF file (source of Link)
		 *	@param	printCopies		Integer with number of copies to print
		 *	@param	parameters		String array of parameter values to execute report
		 *
         */
		private static void sendToPrinter(String printerName , String dataName, String fileName, int printCopies, String[] parameters) throws JRException, Exception
		{
			PrintService printerService = null;
			Map<String,Object> params = new HashMap<String,Object>(); // Storage for rendering parameters.
			JasperPrint jasperPrint = null;
			
			System.out.println("Sending to Printer Mode");
			
			// TODO - Test if XML or SQL
			if (parameters != null) {
				System.out.println("SQL datasource /w Parameters");
				System.out.println("No. of Parameters:" + parameters.length);

				//-----------------------------------------------------
				// Prepare the Parameters to run the report
				
				for (int ctr=0; ctr < parameters.length; ctr++) {
					// Extract the first value pair from the array
					String thisParam = parameters[ctr];

					// Split the value pair
					String[] thisParamArr = thisParam.split("=");

					// Set the value pair in a Hashmap for Jasper
					params.put(thisParamArr[0], thisParamArr[1]);
				}
				
				//-----------------------------------------------------
				// Prepare the Oracle connection

				Connection conn = null;
				try {
					Class.forName ("oracle.jdbc.OracleDriver");

					String url = configProps.getProperty( "jdbc.connection.url" );
					String user = configProps.getProperty( "jdbc.connection.username" );
					String pass = configProps.getProperty( "jdbc.connection.password" );

					conn = DriverManager.getConnection( url , user , pass );

				} catch(SQLException e) {
					System.err.println("ERROR:" + e.getMessage() );				
					throw e;
				}
				
				jasperPrint = JasperFillManager.fillReport( fileName , params , conn);
			} else {
				System.out.println("XML Datasource");

				Document document = JRXmlUtils.parse(new File( dataName ));
				params.put(JRXPathQueryExecuterFactory.PARAMETER_XML_DATA_DOCUMENT, document);
				jasperPrint = JasperFillManager.fillReport( fileName , params );

			}
			
			//list all printers and select the one selected by the user
			PrintRequestAttributeSet aset = new HashPrintRequestAttributeSet();  
			PrintService[] servicesList = PrintServiceLookup.lookupPrintServices(null, aset);
			for(int l=0;l<servicesList.length;l++){
				if(servicesList[l].getName().trim().equals( printerName.trim() )){
					System.out.println("BINGO! Printer queue name:" + servicesList[l].getName() );
					printerService = servicesList[l];
					break;
				}
			}
			
			PrintRequestAttributeSet printRequestAttributeSet = new HashPrintRequestAttributeSet();
			printRequestAttributeSet.add(MediaSizeName.ISO_A4);
			printRequestAttributeSet.add(new Copies( printCopies )); 

			PrintServiceAttributeSet printServiceAttributeSet = new HashPrintServiceAttributeSet();
			printServiceAttributeSet.add(new PrinterName( printerName, null));

			JRPrintServiceExporter exporter = new JRPrintServiceExporter();

			exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
			exporter.setParameter(JRPrintServiceExporterParameter.PRINT_SERVICE, printerService);
			exporter.setParameter(JRPrintServiceExporterParameter.PRINT_REQUEST_ATTRIBUTE_SET, printRequestAttributeSet);
			exporter.setParameter(JRPrintServiceExporterParameter.PRINT_SERVICE_ATTRIBUTE_SET, printServiceAttributeSet);
			exporter.setParameter(JRPrintServiceExporterParameter.DISPLAY_PAGE_DIALOG, Boolean.FALSE);
			exporter.setParameter(JRPrintServiceExporterParameter.DISPLAY_PRINT_DIALOG, Boolean.FALSE);
			
			exporter.exportReport();

			// TODO - Clean this up. Trapping error and throwing it away like it never happened.

			//net.sf.jasperreports.engine.JRException: Error printing report.
			//	at net.sf.jasperreports.engine.print.JRPrinterAWT.printPages(JRPrinterAWT.java:197)
			//	at net.sf.jasperreports.engine.print.JRPrinterAWT.printPages(JRPrinterAWT.java:84)
			//	at net.sf.jasperreports.engine.JasperPrintManager.printPages(JasperPrintManager.java:197)
			//	at net.sf.jasperreports.engine.JasperPrintManager.printReport(JasperPrintManager.java:88)
			//	at BuildReport6.sendToPrinter(BuildReport6.java:589)
			//	at BuildReport6.main(BuildReport6.java:194)
			//Caused by: java.awt.print.PrinterException: No print service found.
			//	at sun.print.RasterPrinterJob.print(Unknown Source)
			//	at sun.print.RasterPrinterJob.print(Unknown Source)
			//	at net.sf.jasperreports.engine.print.JRPrinterAWT.printPages(JRPrinterAWT.java:192)
			//	... 5 more

			try {
				JasperPrintManager.printReport(jasperPrint, false);
			} catch(JRException e) {
				System.err.println("ERROR in sendToPrinter():" + e.getMessage() );				
			} catch(Exception e) {
				System.err.println("Nothing to see here, move along.");
			}
			
		}
		
		private static void usage()
        {
            System.out.println( "This program requires:" );
            System.out.println( "   a) Configuration File" );
            System.out.println( "   b) Command Line arguments\r\n" );

            System.out.println( "(a) The path configuration file must be set with an environment variable CONFIG_FILENAME" );
            System.out.println( "    eg. SET CONFIG_FILENAME=c:\\somewhere\\myconfig.properties" );
            System.out.println( "    The contents must set the following:" );
            System.out.println( "        mail.smtp.host" );
            System.out.println( "        mail.product.datasheet.path" );
            System.out.println( "        jdbc.connection.url" );
            System.out.println( "        jdbc.connection.username" );
            System.out.println( "        jdbc.connection.username\r\n" );
            
            System.out.println( "(b) To execute this program you need the following arguments:" );
			System.out.println( "    java -jar BuildReport.jar -Ttask -Ffile -S -O -P -C -ET -EF -V" );
			System.out.println( "        -T : print | email | pdfprint | pdfemail | pdf" );
			System.out.println( "        -F : file name of compiled report eg. invoice.jasper" );
			System.out.println( "        -S : file name of XML source data eg. 2601234.xml" );
			System.out.println( "        -O : file name of output eg. invoice.pdf" );
			System.out.println( "        -P : printer queue name. Must match windows print queue eg. CADMGEN3" );
			System.out.println( "        -C : number of copies to print eg: 2" );
			System.out.println( "        -ET : Email to address" );
			System.out.println( "        -EF : Email from address" );
			System.out.println( "        -V : Parameters to pass into JRXML template eg. param1=value1\r\n" );

        }
		
        /**
         *	Function to send report output via email
		 *
		 *	@param	mailFromAddr		Text string with email address of sender
		 *	@param	mailToAddr			Text string with email address of recipient
		 *	@param	mailFileAttName		Text string with path + file name to pick up
		 *	@param	multiDoc			W3C document of Source data. This is used to determine the subject line
		 *	@param	templateFilename	Text string with template name. This is used to import XML document
		 *
         */
		private static void sendemail(String mailFromAddr, String mailToAddr, String mailFileAttName, Document document, String templateFilename ) throws Exception
		{

			// Check length of email address - If less than 4, must be Evolution requester
			if( mailToAddr == null || mailToAddr.length() < 4 ) {
				System.err.println("ERROR: No To Email Address Supplied, bailing out" );
				throw new AddressException("No To Email Address Supplied");
			}

			if( mailFromAddr == null || mailFromAddr.length() < 4 ) {
				System.err.println("ERROR: No Email From Address Supplied, bailing out" );
				throw new AddressException("No Email From Address Supplied");
			}
			
			// Get system properties
			Properties properties = System.getProperties(); 

			// Setup mail server
			String host = configProps.getProperty( "mail.smtp.host" );
			properties.setProperty("mail.smtp.host", host);
			
			// Get the default Session object.
			Session session = Session.getDefaultInstance(properties );

			// Create a default MimeMessage object.
			MimeMessage message = new MimeMessage(session);

			// Set the RFC 822 "From" header field using the
			// value of the InternetAddress.getLocalAddress method.
			message.setFrom(new InternetAddress( mailFromAddr ));

			// Add the given addresses to the specified recipient type.
			message.addRecipient(Message.RecipientType.TO, new InternetAddress(mailToAddr));

			//------------------------------------------------------------------------------------
			MimeBodyPart textPart = new MimeBodyPart();
			textPart.setContent("<p>Your file is attached</p><br>", "text/html");

			// Attach the main file
			MimeBodyPart attachFilePart = new MimeBodyPart();		
			FileDataSource fds = new FileDataSource( mailFileAttName );
			attachFilePart.setDataHandler(new DataHandler(fds));
			attachFilePart.setFileName(fds.getName());

			// Start assembling the email message with "mp"
			Multipart mp = new MimeMultipart();
			mp.addBodyPart(textPart);
			mp.addBodyPart(attachFilePart);
						
			//-----------------------------------------------------------------------------------------------

			// Check if it is a Multi-Doc. If so,populate "myDoc" object.
			//TODO Excessive logic. Simplify. just check the templateFilename.
			MultiDoc myDoc = isMultiDoc( document , templateFilename );
			
			// Check Doc type = PO
			if("porder".equals( myDoc.getDocElement() ) ) {
				
				// Get products
				NodeList list = document.getElementsByTagName( "productId" );

				// Iterator for each product in the document
				for(int x = 0 ; x < list.getLength() ; x++)
				{
					// Get the <productId> element from the XML
					Element currentNode = (Element) list.item(x);

					// Extract the contents (part number)
					String productId = currentNode.getFirstChild().getTextContent();

					// Get the full path from the config file
					String fullPath = configProps.getProperty("mail.product.datasheet.path") + productId + ".pdf";

					// Setup stream objects to read in file
					MimeBodyPart attachment = new MimeBodyPart();		
					FileDataSource productSheet = new FileDataSource( fullPath  );

				    // Check if file exists
					if( productSheet.getFile().exists() ) {
						System.out.println("Attaching: " + fullPath );

						attachment.setDataHandler(new DataHandler(productSheet));
						attachment.setFileName(productSheet.getName());
						mp.addBodyPart(attachment);
						
					} else {
						System.out.println("WARNING: Skipping Attachment (does not exist) - " + fullPath );
					}
				}

			}

			//-----------------------------------------------------------------------------------------------
			
			String subjectStr;

			// If its a single document the getDocCount will be 1
			// If the document type is recognised, it will populate getDocElement, getDocRoot and getDocKey properties.
			
			if( myDoc.getDocCount() > 1 || myDoc.getDocRoot() != null){
				if( myDoc.getDocElement().equals( "DESPATCH" ) ) {
					subjectStr = "Despatch document from Pacific NonWovens.";
				} else if( myDoc.getDocElement().equals( "INVOICE" )) {
					subjectStr = "Invoice document from Pacific NonWovens.";
				} else if( myDoc.getDocElement().equals( "STATEMENT" )) {
					subjectStr = "Monthly Statement from Pacific NonWovens.";
				} else if( myDoc.getDocElement().equals( "Remittance" )) {
					subjectStr = "Remittance advice from Pacific NonWovens.";
				} else {
					subjectStr = "Document: " + fds.getName();
				}
			} else {
				subjectStr = "Document: " + fds.getName();
			}

			// Log the results to screen
			System.out.println( "Email Subject is:" + subjectStr );
			
			// Set the "Subject" header field.
			message.setSubject( subjectStr  ); 

			// Add all the text & attachments in
			message.setContent(mp);
				
			// Send message
			Transport.send(message);
            System.out.println("Email sent to:" + mailToAddr);
            System.out.println("Attached File:" + mailFileAttName);
			
		}
		
        /**
         *	Function to lookup a requestor (eg. SYS) and return an email address. This creates an SQL connection or Oracle and queries the MISC table. Note: Hard coded to UB
		 *
		 *	@param	requestor		Text string with requestor
		 *  @return String containing Email address (eg. admin@unitedbonded.com.au)
		 *
         */
		private static String getEmailAddress(String requestor) throws Exception
		{
			System.out.println("Looking up email address for requestor: " + requestor );
			String getEmailAddress = null;
			Class.forName ("oracle.jdbc.OracleDriver");

			//TODO Remove Hard coded "UB" and pass company in
			String sqlStr = "select  miscalpha_20_1 || miscalpha_20_2 as EmailAddress from misc_data where misc1_co_site = 'UB' and misc1_rec_type = 130 and misc1_ref_1 = '" + requestor + "'";

			try {

				String url = configProps.getProperty( "jdbc.connection.url" );
				String user = configProps.getProperty( "jdbc.connection.username" );
				String pass = configProps.getProperty( "jdbc.connection.password" );
				
				Connection conn = DriverManager.getConnection( url , user , pass );

				try {

					Statement stmt = conn.createStatement();
					try {
						ResultSet rset = stmt.executeQuery( sqlStr );
						try {
							while (rset.next() ) {
								getEmailAddress = rset.getString(1);
							}
						}
						finally 
						{
							try { rset.close(); } catch (Exception ignore) {}
						}
					} 
					finally 
					{
						try { stmt.close(); } catch (Exception ignore) {}
					}
				} 
				finally 
				{
					try { conn.close(); } catch (Exception ignore) {}
				}
			} 
			catch (SQLException e) 
			{
				System.err.println("ERROR:" + e.getMessage() );				
				throw e;
			}

			System.out.println("SQL Completed. Email address is: " + getEmailAddress);
			return getEmailAddress;
			
		}
}


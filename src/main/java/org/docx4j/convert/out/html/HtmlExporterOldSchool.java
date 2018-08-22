package org.docx4j.convert.out.html;


import java.io.IOException;
import java.io.StringReader;
import java.util.HashMap;
import java.util.List;

import javax.xml.bind.Unmarshaller;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Templates;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.stream.StreamSource;

import org.apache.log4j.Logger;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.Containerization;
import org.docx4j.convert.out.Converter;
//import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.convert.out.PageBreak;
import org.docx4j.jaxb.Context;
import org.docx4j.model.PropertyResolver;
import org.docx4j.model.SymbolModel.SymbolModelTransformState;
import org.docx4j.model.TransformState;
import org.docx4j.model.properties.Property;
import org.docx4j.model.properties.PropertyFactory;
import org.docx4j.model.properties.paragraph.PBorderBottom;
import org.docx4j.model.properties.paragraph.PBorderTop;
import org.docx4j.model.properties.paragraph.PShading;
import org.docx4j.model.table.TableModel.TableModelTransformState;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.CTShd;
import org.docx4j.wml.PPr;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.PPrBase.NumPr;
import org.w3c.dom.Document;
import org.w3c.dom.DocumentFragment;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.w3c.dom.traversal.NodeIterator;
import org.xml.sax.InputSource;

public class HtmlExporterOldSchool extends  AbstractHtmlExporter {
			
	protected static Logger log = Logger.getLogger(HtmlExporterOldSchool.class);
	
	public static void log(String message ) {
		
		log.info(message);
	}
	
	static Templates xslt;	
	
	/**
	 * Some docx use style Heading 1 for entire paragraphs.
	 * Since we convert Word headings to HTML H1, H2 etc,
	 * and don't apply any additional CSS, this has the 
	 * effect of making those paragraphs bold, large.
	 * To work around this, treat Hn as a normal P if
	 * it is longer than MAX_HEADING_LENGTH.
	 */
	static int MAX_HEADING_LENGTH;
		
	static {
		try {
//            Source xsltSource = new StreamSource(new File("docx/docx2xhtmlOldSchool.xslt"));
			
			Source xsltSource = new StreamSource(
						org.docx4j.utils.ResourceUtils.getResource(
								"org/docx4j/convert/out/html/docx2xhtmlOldSchool.xslt"));
			xslt = XmlUtils.getTransformerTemplate(xsltSource);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (TransformerConfigurationException e) {
			e.printStackTrace();
		}
		
		String lengthStr = ConverterProperties.getProperty("xhtml.maxHeadingLength");
		MAX_HEADING_LENGTH = Integer.parseInt(lengthStr);
		
	}
	
	
	// Implement the interface.  	
	
	public void output(javax.xml.transform.Result result) throws Docx4JException {
		
		if (wmlPackage==null) {
			throw new Docx4JException("Must setWmlPackage");
		}
		
		if (htmlSettings==null) {
			log.debug("Using empty HtmlSettings");
			htmlSettings = new HtmlSettings();			
		}		
		
		try {
			html(wmlPackage, result, htmlSettings);
		} catch (Exception e) {
			throw new Docx4JException("Failed to create HTML output", e);
		}		
		
	}
	
	// End interface
	
	/** Create an html version of the document, using CSS font family
	 *  stacks.  This is appropriate if the HTML is intended for
	 *  viewing in a web browser, rather than an intermediate step
	 *  on the way to generating PDF output (not that docx4j
	 *  supports that approach anymore). 
	 * 
	 * @param result
	 *            The javax.xml.transform.Result object to transform into 
	 * 
	 * */ 
	@Override
	@Deprecated	
	public void html(WordprocessingMLPackage wmlPackage, javax.xml.transform.Result result,
			String imageDirPath) throws Exception {
		
		html(wmlPackage, result, true, imageDirPath);
	}
	
	@Override
	@Deprecated
	public void html(WordprocessingMLPackage wmlPackage, javax.xml.transform.Result result, boolean fontFamilyStack,
			String imageDirPath) throws Exception {
	
		// Prep parameters
		HtmlSettings htmlSettings = new HtmlSettings();
		htmlSettings.setFontFamilyStack(fontFamilyStack);
		
		if (imageDirPath==null) {
			imageDirPath = "";
		}
		htmlSettings.setImageDirPath(imageDirPath);    	
		
		html(wmlPackage, result, htmlSettings);
	}
	
	/** Create an html version of the document. 
	 * 
	 * @param result
	 *            The javax.xml.transform.Result object to transform into 
	 * 
	 * */ 
	@Override
	public void html(WordprocessingMLPackage wmlPackage,
			javax.xml.transform.Result result, HtmlSettings htmlSettings)
			throws Exception {
	
		// Containerization of borders/shading
		MainDocumentPart mdp = wmlPackage.getMainDocumentPart();
		// Don't change the user's Document object; create a tmp one
		org.docx4j.wml.Document tmpDoc = XmlUtils.deepCopy(wmlPackage
				.getMainDocumentPart().getJaxbElement());
		Containerization.groupAdjacentBorders(tmpDoc.getBody());
		PageBreak.movePageBreaks(tmpDoc.getBody());
	
		org.w3c.dom.Document doc = XmlUtils.marshaltoW3CDomDocument(tmpDoc);
	
		// log.debug( XmlUtils.w3CDomNodeToString(doc));
	
		// Prep parameters
		if (htmlSettings == null) {
			htmlSettings = new HtmlSettings();
			// ..Ensure that the font names in the XHTML have been mapped to
			// these matches
			// possibly via an extension function in the XSLT
		}
	
		// Ensure that the imageHandler is set up
		boolean privateImageHandler = false;
		if (htmlSettings.getImageHandler() == null) {
			htmlSettings.setImageHandler(
				new HTMLConversionImageHandler(htmlSettings.getImageDirPath(), 
											   htmlSettings.getImageTargetUri(), 
											   htmlSettings.isImageIncludeUUID()));
			privateImageHandler = true;
		}
		
		if (htmlSettings.getFontMapper() == null) {
			htmlSettings.setFontMapper(wmlPackage.getFontMapper());
			log.debug("FontMapper set.. ");
		}
	
		htmlSettings.setWmlPackage(wmlPackage);
	
		// Allow arbitrary objects to be passed to the converters.
		// The objects are assumed to be specific to a particular converter (eg
		// table),
		// so assume there will be one object implementing TransformState per
		// converter.
		HashMap<String, TransformState> modelStates = new HashMap<String, TransformState>();
		htmlSettings.getSettings().put("modelStates", modelStates);
	
		// Converter c = new Converter();
		Converter.getInstance().registerModelConverter("w:tbl",
				new TableWriter());
		Converter.getInstance().registerModelConverter("w:sym",
				new SymbolWriter());
	
		// By convention, the transform state object is stored by reference to
		// the
		// type of element to which its model applies
		modelStates.put("w:tbl", new TableModelTransformState());
		modelStates.put("w:sym", new SymbolModelTransformState());
	
		// .. although that convention can be bent ..
		modelStates.put("footnoteNumber", new FootnoteState());
		modelStates.put("endnoteNumber", new EndnoteState());
	
		Converter.getInstance().start(wmlPackage);
	
		// Now do the transformation
		log.debug("About to transform...");
		org.docx4j.XmlUtils.transform(doc, xslt, htmlSettings.getSettings(),
				result);
	
		if (privateImageHandler) {
			//remove a locally created imageHandler in case the HtmlSettings get reused
			htmlSettings.getSettings().remove(HtmlSettings.IMAGE_HANDLER);
		}
		log.info("wordDocument transformed to xhtml ..");
	
	}
	
	/* ---------------Modified from AbstractHtmlExporter  ---------------- */
	
//	/**
//	 * Properties (CSS names) which should not appear in the
//	 * HTML output.
//	 */
//	private static List<String> excludedProperties;	
//    public static void setExcludedProperties(List<String> excludedPropertiesList) {
//		excludedProperties = excludedPropertiesList;
//	}
//    
//    private final static String ALL_EXCEPT_BASIC_FONT_STYLES = "ALL_EXCEPT_BASIC_FONT_STYLES";
//
//	public static void createCss(OpcPackage wmlPackage, PPr pPr, StringBuffer result, boolean ignoreBorders) {
//    	
//		if (pPr==null) {
//			return;
//		}
//    	
//    	List<Property> properties = PropertyFactory.createProperties(wmlPackage, pPr);    	
//    	for( Property p :  properties ) {
//
//    		if ( excludedProperties.contains(p.getCssName()) ) {
//    			log.debug("skipping " + p.getCssName() );
//    			continue;
//    		} 
//    		
//			if (ignoreBorders &&
//					((p instanceof PBorderTop)
//							|| (p instanceof PBorderBottom))) {
//				continue;
//			}
//			
//			if (p instanceof PShading) {
//    	    	// To close the gap between divs, we need to avoid
//    	    	// CSS margin collapse.    	    	
//    	    	// To do that, we add a border the same color as 
//    	    	// the background color				
//				String fill = ((CTShd)p.getObject()).getFill();				
//				result.append("border-color: #" + fill + "; border-style:solid; border-width:1px;");
//			}
//    		
//    		result.append(p.getCssProperty());
//    	}    
//    }
	    
    /**
     * Compared to the version in NG2, this run-level createCss 
     * filters out certain properties, and facilitates the
     * use of the following XHTML tags:
     * <!ENTITY % fontstyle.basic "tt | i | b | u | s | strike ">
     * (tt renders text in a teletype or a monospaced font - ignore that
     * one; s and strike are equivalent - we use strike)
     */
    public static FontstyleBasicState createCssForRun(OpcPackage wmlPackage, RPr rPr, StringBuffer result) {

    	List<Property> properties = PropertyFactory.createProperties(wmlPackage, rPr);
    	
    	FontstyleBasicState fontstyles = new FontstyleBasicState();
    	
//    	if (excludedProperties.contains("ALL_EXCEPT_BASIC_FONT_STYLES")) {
    		
    		// We're only interested in: b, i, u
	    	for( Property p :  properties ) {
	    		
    			String cssProp = p.getCssProperty();
    			
    			if (cssProp==null) {
    				log.warn("Unable to get CSS for property " + p.getClass().getName()  );
    				continue;
    			}
    			
    			if (cssProp.equals("font-weight: bold;")) {
    				fontstyles.bold = true;
    			} else if (cssProp.equals("font-style: italic;")) {
    				fontstyles.italic = true;
    			} else if (cssProp.equals("text-decoration: underline;")) {
    				fontstyles.underline = true;
    			} else if (cssProp.equals("text-decoration: line-through;")) {
    				fontstyles.strike = true;
	    		}
	    	}
    		
//    	} else {
//    	
//    		// Be particular about what we are dropping
//	    	for( Property p :  properties ) {
//	    		if ( excludedProperties.contains(p.getCssName()) ) {
//	    			log.debug("skipping " + p.getCssName() );
//	    		} else {
//	    			String cssProp = p.getCssProperty();
//	    			if (cssProp.equals("font-weight: bold;")) {
//	    				fontstyles.bold = true;
//	    			} else if (cssProp.equals("font-style: italic;")) {
//	    				fontstyles.italic = true;
//	    			} else if (cssProp.equals("text-decoration: underline;")) {
//	    				fontstyles.underline = true;
//	    			} else if (cssProp.equals("text-decoration: line-through;")) {
//	    				fontstyles.strike = true;
//		    		} else {
//		    			result.append(p.getCssProperty());
//		    		}
//	    		}
//	    	}
//    	}
    	return fontstyles;
    }
    
    
	/* ---------------Xalan XSLT Extension Functions ---------------- */
	
	
	public static DocumentFragment notImplemented(NodeIterator nodes, String message) {
	
		Node n = nodes.nextNode();
		log.warn("NOT IMPLEMENTED: support for "+ n.getNodeName() + "; " + message);
		
		if (log.isDebugEnabled() ) {
			
			if (message==null) message="";
			
			log.debug( XmlUtils.w3CDomNodeToString(n)  );
	
			// Return something which will show up in the HTML
			return message("NOT IMPLEMENTED: support for " + n.getNodeName() + " - " + message);
		} else {
			
			// Put it in a comment node instead?
			
			return null;
		}
	}
	
	public static DocumentFragment message(String message) {
		
		log.info(message);
		
		if (!log.isDebugEnabled()) return null;
	
		String html = "<div style=\"color:red\" >"
			+ message
			+ "</div>";  
	
		javax.xml.parsers.DocumentBuilderFactory dbf = DocumentBuilderFactory
				.newInstance();
		dbf.setNamespaceAware(true);
		StringReader reader = new StringReader(html);
		InputSource inputSource = new InputSource(reader);
		Document doc = null;
		try {
			doc = dbf.newDocumentBuilder().parse(inputSource);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		reader.close();
	
		DocumentFragment docfrag = doc.createDocumentFragment();
		docfrag.appendChild(doc.getDocumentElement());
		return docfrag;		
	}
			
    
    
    public static DocumentFragment createBlockForSdt( 
    		WordprocessingMLPackage wmlPackage,
    		NodeIterator pPrNodeIt,
    		String pStyleVal, NodeIterator childResults, String tag) {
    	
    	DocumentFragment docfrag = createBlock( wmlPackage,
        		 pPrNodeIt,
        		 pStyleVal,  childResults,
        		 "div", null, null);
    	    	    
    	return docfrag;
    }	    

    /**
     * In this version of this extension function,
     * we wish to apply an @class="tabn", where
     * n is the numbering level.  In order to do 
     * this, we need level Id (not numId, but
     * pass it for completeness).
     */
    public static DocumentFragment createBlockForPPr( 
    		WordprocessingMLPackage wmlPackage,
    		NodeIterator pPrNodeIt,
    		String pStyleVal, NodeIterator childResults, int contentLength,
    		String numId, String levelId) {
    	
    	// <!ENTITY % heading "h1|h2|h3|h4|h5|h6">
    	// in this OldSchool exporter, these are not to be styled.
    	// TODO: foreign language support.
    	// what if the style is basedOn one of these? 
    	
    	if (pStyleVal.startsWith("Heading")
    			&& contentLength>MAX_HEADING_LENGTH) {
    		log.info("Treating long heading as plain p");
        	return createBlock( 
           		 wmlPackage,
           		 pPrNodeIt,
           		 pStyleVal,  childResults,
           		  "p", numId, levelId );    		
    	} else {
    		System.out.println(pStyleVal);
    		System.out.println(contentLength);    		
    	}
    	
    	if (pStyleVal.equals("Heading1")) {
        	return createBlockWithoutCSS( wmlPackage, childResults, "h1" );    		
    	}
    	if (pStyleVal.equals("Heading2")) {
        	return createBlockWithoutCSS( wmlPackage, childResults, "h2" );    		
    	}
    	if (pStyleVal.equals("Heading3")) {
        	return createBlockWithoutCSS( wmlPackage, childResults, "h3" );    		
    	}
    	if (pStyleVal.equals("Heading4")) {
        	return createBlockWithoutCSS( wmlPackage, childResults, "h4" );    		
    	}
    	if (pStyleVal.equals("Heading5")) {
        	return createBlockWithoutCSS( wmlPackage, childResults, "h5" );    		
    	}
    	if (pStyleVal.equals("Heading6")) {
        	return createBlockWithoutCSS( wmlPackage, childResults, "h6" );    		
    	}

    	return createBlock( 
        		 wmlPackage,
        		 pPrNodeIt,
        		 pStyleVal,  childResults,
        		  "p", numId, levelId );
    	
    }
    
    private static DocumentFragment createBlockWithoutCSS( 
    		WordprocessingMLPackage wmlPackage,
    		NodeIterator childResults,
    		String htmlElementName ) {
    	
        try {
        	        	
            // Create a DOM document to take the results			
        	DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();        
			Document document = factory.newDocumentBuilder().newDocument();			
				//log.info("Document: " + document.getClass().getName() );
			Element xhtmlBlock = document.createElement(htmlElementName);			
			document.appendChild(xhtmlBlock);
									    						
			Node n = childResults.nextNode();
			do {	
				
				// getNumberXmlNode creates a span node, which is empty
				// if there is no numbering.
				// Let's get rid of any such <span/>.
				
				// What we actually get is a document node
				if (n.getNodeType()==Node.DOCUMENT_NODE) {
					log.debug("handling DOCUMENT_NODE");
					// Do just enough of the handling here
	                NodeList nodes = n.getChildNodes();
	                if (nodes != null) {
	                    for (int i=0; i<nodes.getLength(); i++) {
	                    	
	        				if (((Node)nodes.item(i)).getLocalName().equals("span")
	        						&& ! ((Node)nodes.item(i)).hasChildNodes() ) {
	        					// ignore
	        					log.debug(".. ignoring <span/> ");
	        				} else {
	        					XmlUtils.treeCopy( (Node)nodes.item(i),  xhtmlBlock );	        					
	        				}
	                    }
	                }					
				} else {
					XmlUtils.treeCopy( n,  xhtmlBlock );
				}
				// next 
				n = childResults.nextNode();
				
			} while ( n !=null ); 
			
			DocumentFragment docfrag = document.createDocumentFragment();
			docfrag.appendChild(document.getDocumentElement());

			return docfrag;
						
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.toString() );
			log.error(e);
		} 
    	
    	return null;
    	
    }

    private static DocumentFragment createBlock( 
    		WordprocessingMLPackage wmlPackage,
    		NodeIterator pPrNodeIt,
    		String pStyleVal, NodeIterator childResults,
    		String htmlElementName,
    		String numId, String levelId) {
    	
		try {
        	
            // Create a DOM document to take the results			
        	DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();        
			Document document = factory.newDocumentBuilder().newDocument();			
				//log.info("Document: " + document.getClass().getName() );
			Element xhtmlBlock = document.createElement(htmlElementName);			
			document.appendChild(xhtmlBlock);
									    
			// Set @class="tabn" where n is list level
			if ( pStyleVal ==null || pStyleVal.equals("") ) {
//				pStyleVal = "Normal";
				pStyleVal = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart().getDefaultParagraphStyle().getStyleId();
			}
			int listLevel = getListLevel(wmlPackage, pStyleVal,  numId,  levelId);
			if (listLevel >=0 ) {
				xhtmlBlock.setAttribute("class", "tab" + listLevel);
			}
						
			Node n = childResults.nextNode();
			do {	
				
				// getNumberXmlNode creates a span node, which is empty
				// if there is no numbering.
				// Let's get rid of any such <span/>.
				
				// What we actually get is a document node
				if (n.getNodeType()==Node.DOCUMENT_NODE) {
					log.debug("handling DOCUMENT_NODE");
					// Do just enough of the handling here
	                NodeList nodes = n.getChildNodes();
	                if (nodes != null) {
	                    for (int i=0; i<nodes.getLength(); i++) {
	                    	
	        				if (((Node)nodes.item(i)).getLocalName().equals("span")
	        						&& ! ((Node)nodes.item(i)).hasChildNodes() ) {
	        					// ignore
	        					log.debug(".. ignoring <span/> ");
	        				} else {
	        					XmlUtils.treeCopy( (Node)nodes.item(i),  xhtmlBlock );	        					
	        				}
	                    }
	                }					
				} else {
					
	//					log.info("Node we are importing: " + n.getClass().getName() );
	//					foBlockElement.appendChild(
	//							document.importNode(n, true) );
					/*
					 * Node we'd like to import is of type org.apache.xml.dtm.ref.DTMNodeProxy
					 * which causes
					 * org.w3c.dom.DOMException: NOT_SUPPORTED_ERR: The implementation does not support the requested type of object or operation.
					 * 
					 * See http://osdir.com/ml/text.xml.xerces-j.devel/2004-04/msg00066.html
					 * 
					 * So instead of importNode, use 
					 */
					XmlUtils.treeCopy( n,  xhtmlBlock );
				}
				// next 
				n = childResults.nextNode();
				
			} while ( n !=null ); 
			
			DocumentFragment docfrag = document.createDocumentFragment();
			docfrag.appendChild(document.getDocumentElement());

			return docfrag;
						
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.toString() );
			log.error(e);
		} 
    	
    	return null;
    	
    }
    
    public static int getListLevel(WordprocessingMLPackage wmlPackage,
    		//NodeIterator pPrNodeIt,
    		String pStyleVal, String numId, String levelId) {
    	
    	// If we're provided with a levelId, things are simple
    	if (levelId!=null
    			&& !levelId.equals("")) {
    		return Integer.parseInt(levelId);
    	}
    	
    	org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart numberingPart =
        		wmlPackage.getMainDocumentPart().getNumberingDefinitionsPart();        	
    	if (numberingPart==null) {
    		return -1;
    	}
    	    	
    	// First, establish whether this is numbered at all
    	// (ie, is there a numId?)
    	
    	// If numId is not provided explicitly, 
    	// is it provided by the style?
    	// (ie does this style have a list associated with it?)
    	if (numId != null 
    			&& !numId.equals("")) {
    		log.debug("Using numId: " + numId);  
    		// but ilvl is implicit
    		return 0;
    	} 

    	// Have to get it from styles
		org.docx4j.wml.Style style = null;
		if (pStyleVal==null || pStyleVal.equals("") ) {
    		log.debug("no explicit numId; no style either");
			return -1;
		}
		
		log.debug("no explicit numId; looking in styles");
		// First, try this style alone
    	PropertyResolver propertyResolver = wmlPackage.getMainDocumentPart().getPropertyResolver();
		style = propertyResolver.getStyle(pStyleVal); 
		
    	if (style == null) {
    		log.debug("Couldn't find style '" + pStyleVal + "'");
    		return -1;
    	} 
    	
    	NumPr numPr;
    	if (style.getPPr() != null) {

    		numPr = style.getPPr().getNumPr();
		
	    	if (numPr!=null) {
				if (numPr.getIlvl()!=null) {
					return numPr.getIlvl().getVal().intValue();
				}        		
				if (numPr.getNumId()!=null) {
					return 0;  // default
				}
				return -1;
	    	}    		
    	}
    	
		// numPr==null
			
    	// Second, use propertyResolver to follow <w:basedOn w:val="blagh"/>
		log.debug(pStyleVal + ".. use propertyResolver to follow basedOn");
		PPr ppr = propertyResolver.getEffectivePPr(pStyleVal);
		numPr = ppr.getNumPr();

		// Same stuff again
    	if (numPr==null) {
    		return -1;
    	} else {
			if (numPr.getIlvl()!=null) {
				return numPr.getIlvl().getVal().intValue();
			}        		
			if (numPr.getNumId()!=null) {
				return 0;  // default
			}
			return -1;
    	}    		    			
    }

    public static DocumentFragment createBlockForRPr( 
    		WordprocessingMLPackage wmlPackage,
    		NodeIterator pPrNodeIt,
    		NodeIterator rPrNodeIt,
    		NodeIterator childResults ) {
    
    	PropertyResolver propertyResolver = 
        		wmlPackage.getMainDocumentPart().getPropertyResolver();
   	    	
    	// Note that this is invoked for every paragraph with a pPr node.
    	
    	// incoming objects are org.apache.xml.dtm.ref.DTMNodeIterator 
    	// which implements org.w3c.dom.traversal.NodeIterator

    	
//    	log.info("rPrNode:" + rPrNodeIt.getClass().getName() ); // org.apache.xml.dtm.ref.DTMNodeIterator    	
//    	log.info("childResults:" + childResults.getClass().getName() ); 
    	
    	
        try {
        	
			Unmarshaller u = Context.jc.createUnmarshaller();			
			u.setEventHandler(new org.docx4j.jaxb.JaxbValidationEventHandler());

			// If there is w:pPr/w:pStyle,			
			// we need to honour any rPr in the pStyle
			PPr pPrDirect = null;
        	if (pPrNodeIt!=null) {
        		Node n = pPrNodeIt.nextNode();
        		if (n!=null) {
        			Object jaxb = u.unmarshal(n);
        			try {
        				pPrDirect =  (PPr)jaxb;
        			} catch (ClassCastException e) {
        		    	log.error("Couldn't cast " + jaxb.getClass().getName() + " to PPr!");
        			}        	        			
        		}
        	}
        	
			Object jaxbR = u.unmarshal(rPrNodeIt.nextNode());			
			RPr rPrDirect = null;
			try {
				rPrDirect =  (RPr)jaxbR;
			} catch (ClassCastException e) {
		    	log.error("Couldn't cast .." );
			}        	
        	RPr rPr = propertyResolver.getEffectiveRPr(rPrDirect, pPrDirect);
        	
            // Create a DOM builder and parse the fragment
        	DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();        
			Document document = factory.newDocumentBuilder().newDocument();
			Node attachmentPoint = document;
			
			//log.info("Document: " + document.getClass().getName() );
							
			if (log.isDebugEnabled()) {					
				log.debug(XmlUtils.marshaltoString(rPr, true, true));					
			}
			
			// Does our rPr contain anything else?
			StringBuffer inlineStyle =  new StringBuffer();
			FontstyleBasicState fontstyles = createCssForRun(wmlPackage, rPr, inlineStyle);				
			if (!inlineStyle.toString().equals("") ) {
				Node span = document.createElement("span");			
				document.appendChild(span);
				((Element)span).setAttribute("style", inlineStyle.toString() );
				attachmentPoint = span;
			}
			
			if (fontstyles.bold) {
				Node b = document.createElement("b");			
				attachmentPoint.appendChild(b);
				attachmentPoint = b;				
			}
			if (fontstyles.italic) {
				Node i = document.createElement("i");			
				attachmentPoint.appendChild(i);
				attachmentPoint = i;				
			}
			if (fontstyles.underline) {
				Node underline = document.createElement("u");			
				attachmentPoint.appendChild(underline);
				attachmentPoint = underline;				
			}
			if (fontstyles.strike) {
				Node strike = document.createElement("strike");			
				attachmentPoint.appendChild(strike);
				attachmentPoint = strike;				
			}

			DocumentFragment docfrag = document.createDocumentFragment();
			
			Node n = childResults.nextNode();
			// with the objective of not adding unnecessary spans:-
			if (attachmentPoint.getNodeName().equals("#document") ) {
				
				if ( n.getChildNodes().getLength()==1 ) {
					if ( n.getFirstChild().getNodeType()==Node.TEXT_NODE) {
						// The special case we have to handle
						// (can't add a text node at document level)
		            	Node textNode = document.createTextNode(n.getFirstChild().getNodeValue());   
		            	docfrag.appendChild(textNode);	
					} else {
						XmlUtils.treeCopy( n,  attachmentPoint );
						docfrag.appendChild(document.getDocumentElement());							
					}
				} else {
					// One of them could be a text node, so wrap in a span
					Node span = document.createElement("span");			
					document.appendChild(span);
					XmlUtils.treeCopy( n,  span );
					docfrag.appendChild(document.getDocumentElement());
				}
				
			} else {
				XmlUtils.treeCopy( n,  attachmentPoint );
				docfrag.appendChild(document.getDocumentElement());
			}
			

			return docfrag;
						
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(e.toString() );
			log.error(e);
		} 
    	
    	return null;
    	
    }
    
    public static int getNextFootnoteNumber(HashMap<String, TransformState> modelStates) {
    	
    	FootnoteState fs = (FootnoteState)modelStates.get("footnoteNumber");
    	return fs.getNextFootnoteNumber();
    }
    
    public static class FootnoteState implements TransformState {
    
	    int footnoteNumber=0;
	    public int getNextFootnoteNumber() {
	    	footnoteNumber++;
	    	return footnoteNumber;
	    	
	    }
    }

    public static int getNextEndnoteNumber(HashMap<String, TransformState> modelStates) {
    	
    	EndnoteState fs = (EndnoteState)modelStates.get("endnoteNumber");
    	return fs.getNextEndnoteNumber();
    }
    
    public static class EndnoteState implements TransformState {
    
	    int endnoteNumber=0;
	    public int getNextEndnoteNumber() {
	    	endnoteNumber++;
	    	return endnoteNumber;
	    	
	    }
    }
    
    public static class FontstyleBasicState {
    	
    	// <!ENTITY % fontstyle.basic "tt | i | b | u | s | strike ">
    	
    	boolean italic;
    	boolean bold;
    	boolean underline;
    	boolean strike;    	        
    }

   
}
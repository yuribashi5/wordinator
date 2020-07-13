/**
 * 
 */
package org.wordinator.xml2docx.generator;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;
import javax.xml.namespace.QName;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFAbstractFootnoteEndnote;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFHeaderFooter;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFNum;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTable.XWPFBorderType;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlCursor.TokenType;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
//import org.openxmlformats.schemas.officeDocument.x2006.math.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtrRef;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
//import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STChapterSep;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STSectionMark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTextDirection;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalAlignRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.STOnOffImpl;
//import org.wordinator.xml2docx.MakeDocx;
import org.wordinator.xml2docx.xwpf.model.XWPFHeaderFooterPolicy;

/**
 * Generates DOCX files from Simple Word Processing Markup Language XML.
 */
public class DocxGenerator {

	/**
	 * Holds a set of table border styles
	 *
	 */
	protected class TableBorderStyles {

		// Default border type is set by the @borderstyle or @framestyle attribute.
		// By default there are no explicit borders.
		XWPFBorderType defaultBorderType = null;
		XWPFBorderType topBorder = null;
		XWPFBorderType bottomBorder = null;
		XWPFBorderType leftBorder = null;
		XWPFBorderType rightBorder = null;
		XWPFBorderType rowSepBorder = null;
		XWPFBorderType colSepBorder = null;

		public TableBorderStyles(XWPFBorderType defaultBorderType, XWPFBorderType topBorder,
				XWPFBorderType bottomBorder, XWPFBorderType leftBorder, XWPFBorderType rightBorder) {

		}

		/**
		 * Construct using specified border styles as the initial values.
		 * 
		 * @param parentBorderStyles Styles to be inherited from parent
		 */
		public TableBorderStyles(TableBorderStyles parentBorderStyles) {
			defaultBorderType = parentBorderStyles.getDefaultBorderType();
			topBorder = parentBorderStyles.getTopBorder();
			bottomBorder = parentBorderStyles.getBottomBorder();
			leftBorder = parentBorderStyles.getLeftBorder();
			rightBorder = parentBorderStyles.getRightBorder();
			rowSepBorder = parentBorderStyles.getRowSepBorder();
			colSepBorder = parentBorderStyles.getColSepBorder();
		}

		/**
		 * Construct initial border styles from an element that may specify border frame
		 * style attributes.
		 * 
		 * @param borderStyleSpecifier XML element that may specify frame style
		 *                             attributes (table, td)
		 */
		public TableBorderStyles(XmlObject borderStyleSpecifier) {

			XmlCursor cursor = borderStyleSpecifier.newCursor();
			String tagname = cursor.getName().getLocalPart();
			String styleValue = null;
			String styleBottomValue = null;
			String styleTopValue = null;
			String styleLeftValue = null;
			String styleRightValue = null;

			if ("table".equals(tagname)) {
				styleValue = cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_ATT);
				styleBottomValue = cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_BOTTOM_ATT);
				styleTopValue = cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_TOP_ATT);
				styleLeftValue = cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_LEFT_ATT);
				styleRightValue = cursor.getAttributeText(DocxConstants.QNAME_FRAMESTYLE_RIGHT_ATT);
			} else {
				styleValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_ATT);
				styleBottomValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_BOTTOM_ATT);
				styleTopValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_TOP_ATT);
				styleLeftValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_LEFT_ATT);
				styleRightValue = cursor.getAttributeText(DocxConstants.QNAME_BORDER_STYLE_RIGHT_ATT);
			}

			if (styleValue != null) {
				setDefaultBorderType(xwpfBorderType(styleValue));
			}
			if (styleTopValue != null) {
				setTopBorder(xwpfBorderType(styleTopValue));
			}
			if (styleBottomValue != null) {
				setBottomBorder(xwpfBorderType(styleBottomValue));
			}
			if (styleLeftValue != null) {
				setLeftBorder(xwpfBorderType(styleLeftValue));
			}
			if (styleRightValue != null) {
				setRightBorder(xwpfBorderType(styleRightValue));
			}
		}

		public XWPFBorderType getDefaultBorderType() {
			return defaultBorderType;
		}

		public void setDefaultBorderType(XWPFBorderType defaultBorderType) {
			this.defaultBorderType = defaultBorderType;

			if (getTopBorder() == null)
				setTopBorder(defaultBorderType);
			if (getBottomBorder() == null)
				setBottomBorder(defaultBorderType);
			if (getLeftBorder() == null)
				setLeftBorder(defaultBorderType);
			if (getRightBorder() == null)
				setRightBorder(defaultBorderType);
		}

		public XWPFBorderType getTopBorder() {
			return topBorder;
		}

		public void setTopBorder(XWPFBorderType topBorder) {
			this.topBorder = topBorder;
		}

		public XWPFBorderType getBottomBorder() {
			return bottomBorder;
		}

		public void setBottomBorder(XWPFBorderType bottomBorder) {
			this.bottomBorder = bottomBorder;
		}

		public XWPFBorderType getLeftBorder() {
			return leftBorder;
		}

		public void setLeftBorder(XWPFBorderType leftBorder) {
			this.leftBorder = leftBorder;
		}

		public XWPFBorderType getRightBorder() {
			return rightBorder;
		}

		public void setRightBorder(XWPFBorderType rightBorder) {
			this.rightBorder = rightBorder;
		}

		public XWPFBorderType getRowSepBorder() {
			return rowSepBorder;
		}

		public void setRowSepBorder(XWPFBorderType rowSepBorder) {
			this.rowSepBorder = rowSepBorder;
		}

		public XWPFBorderType getColSepBorder() {
			return colSepBorder;
		}

		public void setColSepBorder(XWPFBorderType colSepBorder) {
			this.colSepBorder = colSepBorder;
		}

		public STBorder.Enum getBottomBorderEnum() {
			return getBorderEnumForType(getBottomBorder());
		}

		public STBorder.Enum getTopBorderEnum() {
			return getBorderEnumForType(getTopBorder());
		}

		public STBorder.Enum getLeftBorderEnum() {
			return getBorderEnumForType(getLeftBorder());
		}

		public STBorder.Enum getRightBorderEnum() {
			return getBorderEnumForType(getRightBorder());
		}

		public STBorder.Enum getBorderEnumForType(XWPFBorderType type) {
			STBorder.Enum result = null;
			if (type != null) {
				result = stBorderType(type);
			}
			return result;
		}

		/**
		 * Determine if any borders are explicitly set
		 * 
		 * @return True if one or more borders have a defined style.
		 */
		public boolean hasBorders() {
			boolean result = getDefaultBorderType() != null || getBottomBorder() != null || getTopBorder() != null
					|| getLeftBorder() != null || getRightBorder() != null;
			return result;
		}

	}

	public static final Logger log = LogManager.getLogger(DocxGenerator.class.getSimpleName());

	private File outFile;
	private int dotsPerInch = 72; /* DPI */
	// Map of source IDs to internal object IDs.
	private Map<String, BigInteger> bookmarkIdToIdMap = new HashMap<String, BigInteger>();
	private int idCtr = 0;
	private File inFile;
	private XWPFDocument templateDoc;

	/**
	 * 
	 * @param inFile      File representing input document.
	 * @param outFile     File to write DOCX result to
	 * @param templateDoc DOTX template to initialize result DOCX with (provides
	 *                    style definitions)
	 * @throws Exception             Exception from loading the template document
	 * @throws FileNotFoundException If the template document is not found
	 */
	public DocxGenerator(File inFile, File outFile, XWPFDocument templateDoc) throws FileNotFoundException, Exception {
		this.inFile = inFile;
		this.outFile = outFile;
		this.templateDoc = templateDoc;
	}

	/**
	 * Generate the DOCX file from the input Simple WP ML document.
	 * 
	 * @param xml The XmlObject that holds the Simple WP XML content
	 */
	public void generate(XmlObject xml) throws DocxGenerationException, XmlException, IOException {

		XWPFDocument doc = new XWPFDocument();

		setupNumbering(doc, this.templateDoc);
		setupStyles(doc, this.templateDoc);
		constructDoc(doc, xml);

		FileOutputStream out = new FileOutputStream(outFile);
		doc.write(out);
		doc.close();
	}

	/**
	 * Walk the XML document to create the Word document.
	 * 
	 * @param doc Word document to write to
	 * @param xml Simple ML doc to walk
	 */
	private void constructDoc(XWPFDocument doc, XmlObject xml) throws DocxGenerationException {

		//log.info("+ DocxGenerator-constructDoc() BEGIN...");
		XmlCursor cursor = xml.newCursor();
		//log.info("+ DocxGenerator-constructDoc() BEFORE toFirstChild()");
		cursor.toFirstChild(); // Put us on the root element of the document
		//log.info("+ DocxGenerator-constructDoc() BEFORE push()");
		cursor.push();
		//log.info("+ DocxGenerator-constructDoc() BEFORE pageSequenceProperties");
		XmlObject pageSequenceProperties = null;

		//log.info("+ DocxGenerator-constructDoc() BEFORE toChild(page-sequence-properties)");
		
		if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-sequence-properties"))) {
			// Set up document-level headers. These will apply to the whole
			// document if there are no sections, or to the last section if
			// there are sections. Results in a w:sectPr as the last child
			// of w:body.
			//log.info("+ DocxGenerator-constructDoc()-IF-page-sequence-properties BEFORE setupPageSequence()");
			setupPageSequence(doc, cursor.getObject());
			//log.info("+ DocxGenerator-constructDoc()-IF-page-sequence-properties BEFORE pageSequenceProperties=");
			pageSequenceProperties = cursor.getObject();
		}

		cursor.pop();

		//log.info("+ DocxGenerator-constructDoc() BEFORE cursor to 'body'");
		cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "body"));
		//log.info("+ DocxGenerator-constructDoc() AFTER cursor to 'body'");

		//log.info("+ DocxGenerator-constructDoc() BEFORE handleBody()");
		handleBody(doc, cursor.getObject(), pageSequenceProperties);

	}

	/**
	 * Process the elements in &lt;body&gt;
	 * 
	 * @param doc                    Document to add paragraphs to.
	 * @param xml                    Body element
	 * @param pageSequenceProperties Document-level page sequence properties. Used
	 *                               if there are no section-level page sequence
	 *                               properties.
	 * @return Last paragraph of the body (if any)
	 * @throws DocxGenerationException
	 */
	private XWPFParagraph handleBody(XWPFDocument doc, XmlObject xml, XmlObject pageSequenceProperties)
			throws DocxGenerationException {

		if (log.isDebugEnabled()) {
			//log.debug("handleBody(): starting...");
		}

		XmlCursor cursor = xml.newCursor();

		if (cursor.toFirstChild()) {
			do {
				
				//log.debug("+ [debug BEGIN handleBody Do-loop]");
				
				String tagName = cursor.getName().getLocalPart();
				String namespace = cursor.getName().getNamespaceURI();

				String htmlstyle = null;
				String pagebreak = null;
					
				if ("p".equals(tagName)) {
					XWPFParagraph p = doc.createParagraph();
					htmlstyle = cursor.getAttributeText(DocxConstants.QNAME_HTMLSTYLE_ATT);
					pagebreak = cursor.getAttributeText(DocxConstants.QNAME_PAGEBREAK_ATT);
					
					if(!StringUtils.isEmpty(cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT))) {
						htmlstyle.concat("; " + cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT));
					}				
					
					Map<String, String> mapParaAdditionalParameters = createMapHtmlStyle(htmlstyle, pagebreak);
					mapParaAdditionalParameters = cleanupMapEntries(mapParaAdditionalParameters);
					makeParagraph(p, cursor, mapParaAdditionalParameters);

				} else if ("section".equals(tagName)) {
					handleSection(doc, cursor.getObject(), pageSequenceProperties);

				} else if ("table".equals(tagName)) {
					htmlstyle = cursor.getAttributeText(DocxConstants.QNAME_HTMLSTYLE_ATT);
					pagebreak = cursor.getAttributeText(DocxConstants.QNAME_PAGEBREAK_ATT);
					
					if(!StringUtils.isEmpty(cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT))) {
						htmlstyle.concat("; " + cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT));
					}
					
					Map<String, String> mapTableAdditionalParameters = createMapHtmlStyle(htmlstyle, pagebreak);
					mapTableAdditionalParameters = cleanupMapEntries(mapTableAdditionalParameters);
					
					// rotatepg...
					if(!StringUtils.isEmpty(mapTableAdditionalParameters.values().toString()) 
							&& "true".equals(mapTableAdditionalParameters.get("rotatepg"))
						) {
						CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
						CTPageSz pageSz = sectPr.addNewPgSz();
						pageSz.setH(BigInteger.valueOf(12240)); //12240 Twips = 12240/20 = 612 pt = 612/72 = 8.5"
						pageSz.setW(BigInteger.valueOf(15840)); //15840 Twips = 15840/20 = 792 pt = 792/72 = 11"

					}
					
					// pagebreak...
					if(!StringUtils.isEmpty(pagebreak)) {
						Boolean bPageBreak = Boolean.valueOf(pagebreak);
						XWPFParagraph para = doc.createParagraph();
						para.setPageBreak(bPageBreak);
						para.setSpacingAfterLines(0);
						para.setStyle("PageBreakB4Table");
					}
	
					XWPFTable table = doc.createTable();
					makeTable(table, cursor.getObject(), mapTableAdditionalParameters);

				} else if ("object".equals(tagName)) {
					// - - - - - - - - - - - - - - - - - - - - - - -
					// FIXME: This is currently unimplemented.
					// - - - - - - - - - - - - - - - - - - - - - - -
					makeObject(doc, cursor);

				} else {
					log.warn("handleBody(): Unexpected element {" + namespace + "}:'" + tagName
							+ "' in <body>. Ignored.");
				}
			} while (cursor.toNextSibling());

		}
		// The section properties always go on an empty paragraph.
		XWPFParagraph lastPara = doc.createParagraph();
		lastPara.setSpacingBefore(0);
		lastPara.setSpacingAfter(0);
		return lastPara;
	}

	/**
	 * Handle a &lt;section&gt; element
	 * 
	 * @param doc                       Document we're adding to
	 * @param xml                       &lt;section&gt; element
	 * @param docPageSequenceProperties Document-level page sequence properties
	 */
	private void handleSection(XWPFDocument doc, XmlObject xml, XmlObject docPageSequenceProperties)
			throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();

		XmlObject localPageSequenceProperties = null;

		cursor.push();
		if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-sequence-properties"))) {
			localPageSequenceProperties = cursor.getObject();
		}
		cursor.pop();

		if (localPageSequenceProperties == null) {
			localPageSequenceProperties = docPageSequenceProperties;
		}

		cursor.push();
		cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "body"));
		XWPFParagraph lastPara = handleBody(doc, cursor.getObject(), localPageSequenceProperties);
		cursor.pop();

		/*
		 * if (log.isDebugEnabled()) {
		 * log.debug("handleSection(): Setting sectPr on last paragraph."); }
		 */
		CTPPr ppr = (lastPara.getCTP().isSetPPr() ? lastPara.getCTP().getPPr() : lastPara.getCTP().addNewPPr());
		CTSectPr sectPr = ppr.addNewSectPr();

		String sectionType = cursor.getAttributeText(DocxConstants.QNAME_TYPE_ATT);

		if (sectionType != null) {
			CTSectType type = sectPr.addNewType();
			type.setVal(STSectionMark.Enum.forString(sectionType));
		}

		setupPageSequence(doc, localPageSequenceProperties, sectPr);

		ppr.setSectPr(sectPr);

	}

	/**
	 * Set up a page sequence for a section, as opposed to for the document as a
	 * whole.
	 * 
	 * @param doc    Document
	 * @param object The page-sequence-properties element
	 * @param sectPr The sectPr object to set the page sequence properties on.
	 * @throws DocxGenerationException
	 */
	private void setupPageSequence(XWPFDocument doc, XmlObject xml, CTSectPr sectPr) throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();

		setPageNumberProperties(cursor, sectPr);

		cursor.push();
		if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "headers-and-footers"))) {
			constructHeadersAndFooters(doc, cursor.getObject(), sectPr);
		}
		cursor.pop();
		cursor.push();
		if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-size"))) {
			setPageSize(cursor, sectPr);
		}
		cursor.pop();
	}

	private void setPageSize(XmlCursor cursor, CTSectPr sectPr) {
		CTPageSz pageSize = (sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz());
		String codeValue = cursor.getAttributeText(DocxConstants.QNAME_CODE_ATT);
		if (codeValue != null) {
			try {
				long code = Long.parseLong(codeValue);
				pageSize.setCode(BigInteger.valueOf(code));
			} catch (Exception e) {
				log.warn("setPageSize(): Value \"" + codeValue + " for attribute \"code\" is not a decimal number");
			}
		}
		String orientValue = cursor.getAttributeText(DocxConstants.QNAME_ORIENT_ATT);
		if (orientValue != null) {
			pageSize.setOrient(STPageOrientation.Enum.forString(orientValue));
		}
		String widthVal = cursor.getAttributeText(DocxConstants.QNAME_WIDTH_ATT);
		if (null != widthVal) {
			try {
				long width = Measurement.toTwips(widthVal, getDotsPerInch());
				pageSize.setW(BigInteger.valueOf(width));
			} catch (MeasurementException e) {
				log.warn("setPageSize(): Value \"" + widthVal
						+ " for attribute \"width\" cannot be converted to a twips value");
			}
		}

		String heightVal = cursor.getAttributeText(DocxConstants.QNAME_HEIGHT_ATT);
		if (null != heightVal) {
			try {
				long height = Measurement.toTwips(heightVal, getDotsPerInch());
				pageSize.setH(BigInteger.valueOf(height));
			} catch (MeasurementException e) {
				log.warn("setPageSize(): Value \"" + heightVal
						+ " for attribute \"height\" cannot be converted to a twips value");
			}
		}
	}

	/**
	 * Set up page sequence properties for the entire document, including page
	 * geometry, numbering, and headers and footers.
	 * 
	 * @param doc    Document to be constructed
	 * @param xml    page-sequence-properties element
	 * @param sectPr Section properties to store the page sequence details on.
	 * @throws DocxGenerationException
	 */
	private void setupPageSequence(XWPFDocument doc, XmlObject xml) throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();

		CTDocument1 document = doc.getDocument();
		CTBody body = (document.isSetBody() ? document.getBody() : document.addNewBody());
		CTSectPr sectPr = (body.isSetSectPr() ? body.getSectPr() : body.addNewSectPr());

		setPageNumberProperties(cursor, sectPr);
		cursor.push();

		//log.info("+ DocxGenerator-setupPageSequence() BEFORE headers-and-footers");
		if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "headers-and-footers"))) {
			constructHeadersAndFooters(doc, cursor.getObject());
		}
		
		cursor.pop();
		cursor.push();

		//log.info("+ DocxGenerator-setupPageSequence() BEFORE page-size");
		if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-size"))) {
			setPageSize(cursor, sectPr);
		}
		cursor.pop();

	}

	
	private void setPageNumberProperties(XmlCursor cursor, CTSectPr sectPr) {
		cursor.push();
		if (cursor.toChild(new QName(DocxConstants.SIMPLE_WP_NS, "page-number-properties"))) {
			String start = cursor.getAttributeText(DocxConstants.QNAME_START_ATT);
			String format = cursor.getAttributeText(DocxConstants.QNAME_FORMAT_ATT);
			String chapterSep = cursor.getAttributeText(DocxConstants.QNAME_CHAPTER_SEPARATOR_ATT);
			String chapterStyle = cursor.getAttributeText(DocxConstants.QNAME_CHAPTER_STYLE_ATT);

			if (null != format || null != chapterSep || null != chapterStyle || null != start) {
				CTPageNumber pageNumber = (sectPr.isSetPgNumType() ? sectPr.getPgNumType() : sectPr.addNewPgNumType());
				if (null != format) {
					if ("custom".equals(format)) {
						// FIXME: 
						// Implement translation from XSLT number format values to the equivalent
						// Word number formatting values.
						log.warn("Page number format \"" + format
								+ "\" not supported. Use Word-specific values. Using \"decimal\"");
						format = "decimal";
					}
					STNumberFormat.Enum fmt = STNumberFormat.Enum.forString(format);
					if (fmt != null) {
						pageNumber.setFmt(fmt);
					}
				}
				if (chapterSep != null) {
					STChapterSep.Enum sep = STChapterSep.Enum.forString(chapterSep);
					if (sep != null) {
						pageNumber.setChapSep(sep);
					}
				}
				if (chapterStyle != null) {
					try {
						long val = Long.valueOf(chapterStyle);
						pageNumber.setChapStyle(BigInteger.valueOf(val));
					} catch (NumberFormatException e) {
						log.warn("Value \"" + chapterStyle + "\" of @chapter-style attribute is not an integer.");
					}
				}
				if (start != null) {
					try {
						long val = Long.valueOf(start);
						pageNumber.setStart(BigInteger.valueOf(val));
					} catch (NumberFormatException e) {
						log.warn("Value \"" + start + "\" of @start attribute is not an integer.");
					}
				}
			}
		}
		cursor.pop();
	}

	/**
	 * Construct headers and footers on the document. If there are no sections, this
	 * also sets the headers and footers for the document (which acts as a single
	 * section), otherwise, each section must also create the appropriate header
	 * references.
	 * 
	 * @param doc Document to add headers and footers to.
	 * @param xml headers-and-footers element
	 * @throws DocxGenerationException
	 */
	private void constructHeadersAndFooters(XWPFDocument doc, XmlObject xml) throws DocxGenerationException {
		//log.info("+ DocxGenerator-...-constructHeadersAndFooters(doc, xml");
		constructHeadersAndFooters(doc, xml, null);
	}

	/**
	 * Construct headers and footers on the document. If there are no sections, this
	 * also sets the headers and footers for the document (which acts as a single
	 * section), otherwise, each section must also create the appropriate header
	 * references.
	 * 
	 * @param doc    Document to add headers and footers to.
	 * @param xml    headers-and-footers element
	 * @param sectPr Section properties to add header and footer references to. May
	 *               be null
	 * @throws DocxGenerationException
	 */
	private void constructHeadersAndFooters(XWPFDocument doc, XmlObject xml, CTSectPr sectPr)
			throws DocxGenerationException {
		//log.info("+ DocxGenerator-...-constructHeadersAndFooters(doc, xml, sectPr");
		XmlCursor cursor = xml.newCursor();

		boolean haveOddHeader = false;
		boolean haveEvenHeader = false;
		boolean haveOddFooter = false;
		boolean haveEvenFooter = false;

		boolean isDocument = sectPr == null;

		if (cursor.toFirstChild()) {
			XWPFHeaderFooterPolicy sectionHfPolicy = null;
			if (!isDocument) {
				sectionHfPolicy = new XWPFHeaderFooterPolicy(doc, sectPr);
			}

			do {

				//log.info("+ DocxGenerator-...-constructHeadersAndFooters(doc, xml, sectPr)- do...");
				String tagName = cursor.getName().getLocalPart();
				String namespace = cursor.getName().getNamespaceURI();
				List<CTHdrFtrRef> refs = null;

				if ("header".equals(tagName)) {
					//log.info("+ DocxGenerator-...-constructHeadersAndFooters(doc, xml, sectPr)- do...'header'");
					HeaderFooterType type = getHeaderFooterType(cursor);

					if (type == HeaderFooterType.FIRST) {
						CTSectPr localSectPr = sectPr;
						if (localSectPr == null) {
							// FIXME: Can body be null at this time?
							localSectPr = doc.getDocument().getBody().getSectPr();
						}

						CTOnOff titlePg = (localSectPr.isSetTitlePg() ? localSectPr.getTitlePg()
								: localSectPr.addNewTitlePg());
						titlePg.setVal(STOnOff.TRUE);
					}

					if (type == HeaderFooterType.DEFAULT) {
						haveOddHeader = true;
					}

					if (type == HeaderFooterType.EVEN) {
						haveEvenHeader = true;
					}

					if (isDocument) {
						//log.info("+ DocxGenerator-...-constructHeadersAndFooters()- do...'header' - isDocument");
						// Make document-level header
						XWPFHeader header = doc.createHeader(type);
						makeHeaderFooter(header, cursor.getObject());
					} else {
						//log.info("+ DocxGenerator-...-constructHeadersAndFooters()- do...'header' - !isDocument");
						XWPFHeader header = sectionHfPolicy.createHeader(getSTHFTypeForXWPFHFType(type));
						makeHeaderFooter(header, cursor.getObject());
						refs = sectPr.getHeaderReferenceList();
						CTHdrFtrRef ref = getHeadeFooterRefForType(sectPr, refs, type);
						ref.setId(doc.getRelationId(header.getPart()));
						setHeaderFooterRefType(type, ref);
					}
				} else if ("footer".equals(tagName)) {
					//log.info("+ DocxGenerator-...-constructHeadersAndFooters(doc, xml, sectPr)- do...'footer'");
					HeaderFooterType type = getHeaderFooterType(cursor);

					if (type == HeaderFooterType.DEFAULT) {
						haveOddFooter = true;
					}

					if (type == HeaderFooterType.EVEN) {
						haveEvenFooter = true;
					}

					if (type == HeaderFooterType.FIRST) {
						CTSectPr localSectPr = sectPr;

						if (localSectPr == null) {
							// FIXME: Can body be null at this time?
							localSectPr = doc.getDocument().getBody().getSectPr();
						}

						CTOnOff titlePg = (localSectPr.isSetTitlePg() ? localSectPr.getTitlePg()
								: localSectPr.addNewTitlePg());
						titlePg.setVal(STOnOff.TRUE);
					}

					if (isDocument) {
						// Document-level footer
						XWPFFooter footer = doc.createFooter(type);
						makeHeaderFooter(footer, cursor.getObject());
					} else {
						XWPFFooter footer = sectionHfPolicy.createFooter(getSTHFTypeForXWPFHFType(type));
						makeHeaderFooter(footer, cursor.getObject());
						refs = sectPr.getFooterReferenceList();
						CTHdrFtrRef ref = getHeadeFooterRefForType(sectPr, refs, type);
						ref.setId(doc.getRelationId(footer.getPart()));
						setHeaderFooterRefType(type, ref);
					}

				} else {
					log.warn("constructHeadersAndFooters(): Unexpected element {" + namespace + "}:" + tagName
							+ " in <headers-and-footers>. Ignored.");
				}

			} while (cursor.toNextSibling());
			if (!isDocument) {
				// setDefaultSectionHeadersAndFooters(doc, sectPr, sectionHfPolicy);
			}
			// Now set any default headers and footers from the document:
		}

		if ((haveOddHeader || haveOddFooter) && (haveEvenHeader || haveEvenFooter)) {
			doc.setEvenAndOddHeadings(true);
		}

	}

	private CTHdrFtrRef getHeadeFooterRefForType(CTSectPr sectPr, List<CTHdrFtrRef> refs, HeaderFooterType type) {
		CTHdrFtrRef ref = null;
		STHdrFtr.Enum stType = getSTHFTypeForXWPFHFType(type);
		for (CTHdrFtrRef cand : refs) {
			if (cand.getType() == stType) {
				ref = cand;
				break;
			}
		}
		if (ref == null) {
			ref = sectPr.addNewHeaderReference();
		}
		return ref;
	}

	/**
	 * Sets the default headers and footers for a section, creating references to
	 * the document's headers and footers, if any, for any header on the document
	 * but not already set on the section.
	 * 
	 * @param doc             Document containing the section
	 * @param sectPr          Section properties for the section to set the headers
	 *                        on.
	 * @param sectionHfPolicy The section header/footer policy that holds any
	 *                        headers set on th esection.
	 */
	public void setDefaultSectionHeadersAndFooters(XWPFDocument doc, CTSectPr sectPr,
			XWPFHeaderFooterPolicy sectionHfPolicy) {
		XWPFHeaderFooterPolicy docHfPolicy = new XWPFHeaderFooterPolicy(doc);
		if (docHfPolicy != null) {
			XWPFHeader header = null;
			XWPFFooter footer = null;
			// Default header:
			header = docHfPolicy.getDefaultHeader();
			if (sectionHfPolicy.getDefaultHeader() == null && header != null) {
				CTHdrFtrRef ref = sectPr.addNewHeaderReference();
				ref.setId(doc.getRelationId(header.getPart()));
				ref.setType(STHdrFtr.DEFAULT);
			}
			// Even header:
			header = docHfPolicy.getEvenPageHeader();
			if (sectionHfPolicy.getEvenPageHeader() == null && header != null) {
				CTHdrFtrRef ref = sectPr.addNewHeaderReference();
				ref.setId(doc.getRelationId(header.getPart()));
				ref.setType(STHdrFtr.EVEN);
			}
			// First header:
			header = docHfPolicy.getFirstPageHeader();
			if (sectionHfPolicy.getFirstPageHeader() == null && header != null) {
				CTHdrFtrRef ref = sectPr.addNewHeaderReference();
				ref.setId(doc.getRelationId(header.getPart()));
				ref.setType(STHdrFtr.FIRST);
			}
			footer = docHfPolicy.getDefaultFooter();
			if (sectionHfPolicy.getDefaultFooter() == null && footer != null) {
				CTHdrFtrRef ref = sectPr.addNewFooterReference();
				ref.setId(doc.getRelationId(footer.getPart()));
				ref.setType(STHdrFtr.DEFAULT);
			}
			// Even footer:
			footer = docHfPolicy.getEvenPageFooter();
			if (sectionHfPolicy.getEvenPageFooter() == null && footer != null) {
				CTHdrFtrRef ref = sectPr.addNewFooterReference();
				ref.setId(doc.getRelationId(footer.getPart()));
				ref.setType(STHdrFtr.EVEN);
			}
			// First footer:
			footer = docHfPolicy.getFirstPageFooter();
			if (sectionHfPolicy.getFirstPageFooter() == null && footer != null) {
				CTHdrFtrRef ref = sectPr.addNewFooterReference();
				ref.setId(doc.getRelationId(footer.getPart()));
				ref.setType(STHdrFtr.FIRST);
			}
		}
	}

	public STHdrFtr.Enum getSTHFTypeForXWPFHFType(HeaderFooterType type) {
		switch (type) {
		case EVEN:
			return STHdrFtr.EVEN;
		case FIRST:
			return STHdrFtr.FIRST;
		default:
			return STHdrFtr.DEFAULT;
		}

	}

	public void setHeaderFooterRefType(HeaderFooterType type, CTHdrFtrRef ref) {
		ref.setType(getSTHFTypeForXWPFHFType(type));
	}

	
	/**
	 * Returns a map from the contents of the @htmlstring and @pagebreak attributes
	 */
	private Map<String, String> createMapHtmlStyle(String htmlstyle, String pagebreak) {		
		Map<String, String> mapHtmlStyle = new HashMap<String, String>();

		if (!StringUtils.isEmpty(htmlstyle)) {
			mapHtmlStyle.put("htmlstyle", htmlstyle);
		}

		if (!StringUtils.isEmpty(pagebreak)) {
			mapHtmlStyle.put("pagebreak", pagebreak);
		}

		if(null != mapHtmlStyle) {
			mapHtmlStyle = cleanupMapEntries(mapHtmlStyle);
		}

		return mapHtmlStyle;
	}

	
	/**
	 * Construct the content of a page header or footer
	 * 
	 * @param headerFooter {@link XPWFHeader} or {@link XWPFFooter} to add content
	 *                     to
	 * @param xml          The &lt;header&gt; or &lt;footer&gt; element to process
	 * @throws DocxGenerationException
	 */
	private void makeHeaderFooter(XWPFHeaderFooter headerFooter, XmlObject xml) throws DocxGenerationException {
		XmlCursor cursor = xml.newCursor();
		
		String htmlstyle = null;
		String pagebreak = null;

		if (cursor.toFirstChild()) {
			do {

				//log.info("+ DocxGenerator-...-makeHeaderFooter()- do...");
				String tagName = cursor.getName().getLocalPart();
				String namespace = cursor.getName().getNamespaceURI();

				if ("p".equals(tagName)) {
					XWPFParagraph p = headerFooter.createParagraph();

					htmlstyle = cursor.getAttributeText(DocxConstants.QNAME_HTMLSTYLE_ATT);
					pagebreak = cursor.getAttributeText(DocxConstants.QNAME_PAGEBREAK_ATT);
					
					if(!StringUtils.isEmpty(cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT))) {
						htmlstyle.concat("; " + cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT));
					}
					
					Map<String, String> mapHFAdditionalParameters = createMapHtmlStyle(htmlstyle, pagebreak);
					makeParagraph(p, cursor, mapHFAdditionalParameters);

				} else if ("table".equals(tagName)) {
					log.info("+ DocxGenerator-...-makeHeaderFooter()- do...'table'");
					XWPFTable table = headerFooter.createTable(0, 0);
					
					htmlstyle = cursor.getAttributeText(DocxConstants.QNAME_HTMLSTYLE_ATT);
					pagebreak = cursor.getAttributeText(DocxConstants.QNAME_PAGEBREAK_ATT);
					
					if(!StringUtils.isEmpty(cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT))) {
						htmlstyle.concat("; " + cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT));
					}				
					
					Map<String, String> mapTableAdditionalParameters = createMapHtmlStyle(htmlstyle, pagebreak);
					makeTable(table, cursor.getObject(), mapTableAdditionalParameters);
					
				} else {
					// There are other body-level things that could go in a footnote but
					// we aren't worrying about them for now.
					log.warn("makeFootnote(): Unexpected element {" + namespace + "}:" + tagName
							+ "' in <fn>. Ignored.");
				}
			} while (cursor.toNextSibling());
		}
	}

	/**
	 * Get the header or footer type for the element at the cursor.
	 * 
	 * @param cursor
	 * @return {@link HeaderFooterType}
	 */
	private HeaderFooterType getHeaderFooterType(XmlCursor cursor) {
		HeaderFooterType type = HeaderFooterType.DEFAULT;
		String typeName = cursor.getAttributeText(DocxConstants.QNAME_TYPE_ATT);
		if ("even".equals(typeName)) {
			type = HeaderFooterType.EVEN;
		}
		if ("first".equals(typeName)) {
			type = HeaderFooterType.FIRST;
		}
		return type;
	}

	/**
	 * Construct a Word paragraph
	 * 
	 * @param para                 The Word paragraph to construct
	 * @param cursor               Cursor pointing at the
	 *                             <p>
	 *                             element the paragraph will reflect.
	 * @param additionalProperties Additional properties to add to the paragraph,
	 *                             i.e., from sections
	 * @return Paragraph (should be same object as passed in).
	 */
	private XWPFParagraph makeParagraph(XWPFParagraph para, XmlCursor cursor, Map<String, String> mapParaProperties)
			throws DocxGenerationException {

		cursor.push();
		String styleName = cursor.getAttributeText(DocxConstants.QNAME_STYLE_ATT);
		String styleId = cursor.getAttributeText(DocxConstants.QNAME_STYLEID_ATT);

		mapParaProperties = cleanupMapEntries(mapParaProperties);
		
		if (null != styleName && null == styleId) {
			// Look up the style by name:
			XWPFStyle style = para.getDocument().getStyles().getStyleWithName(styleName);
			if (null != style) {
				styleId = style.getStyleId();
			} else {
				// Issue 23: see if this is a latent style and report it
				//
				// This will require an enhancement to the POI API as there is no easy
				// way to get the list of latent styles except to parse out the XML,
				// which I'm not going to--better to fix POI.
				// Unfortunately, there does not appear to be a documented or reliable
				// way to go from Word-defined latent style names to the actual style ID
				// of the style Word *will create* by some internal magic. In addition,
				// any such mapping varies by Word version, locale, etc.
				//
				// That means that in order to use any style it must exist as a proper
				// style.
			}
		}

		if (null != styleId) {
			para.setStyle(styleId);
		}

		// NOTE: renamed additionalProperties to mapParaProperties
		// if (null != mapParaProperties) {

		if (mapParaProperties.containsKey("pagebreak")) {
			Boolean bPageBreak = Boolean.valueOf(mapParaProperties.get("pagebreak"));
			para.setPageBreak(bPageBreak);
		}

		/* Eliot's (as he says) hack... */
		/*
		 * for (String propName : mapParaProperties.keySet()) { String value =
		 * mapParaProperties.get(propName); if (value != null) { // FIXME: This is a
		 * quick hack. Need a more general // and elegant way to manage setting of
		 * properties. if (DocxConstants.PROPERTY_PAGEBREAK.equals(propName)) { if
		 * (DocxConstants.PROPERTY_VALUE_CONTINUOUS.equals(value)) {
		 * para.setPageBreak(false); } else { para.setPageBreak(true); } } } }
		 */
		
//		Integer i = 1;
//		for (Map.Entry<String, String> entry : mapParaProperties.entrySet()) {
//			log.info("+ (" + i++ + ") [makeParagraph dump map entry]: " + entry.toString());
//		}

		// } else {
		// log.error("[makeParagraph...mapParaProperties]: Unexpected mapParaProperties
		// is null]");
		// }

		// Explicit (@page-break-before) page break on a paragraph should override the
		// section-level break I
		// would think.
		String pageBreakBefore = cursor.getAttributeText(DocxConstants.QNAME_PAGE_BREAK_BEFORE_ATT);

		if (!StringUtils.isEmpty(pageBreakBefore)) {
			boolean breakValue = Boolean.valueOf(pageBreakBefore);
			para.setPageBreak(breakValue);
		}
		
		Map<String, String> mapRunProperties = mapParaProperties;

		if (mapRunProperties.containsValue("font-size")) {
			log.debug("+ [debug makeParagraph() BUILT mapRunProperties]: " + mapRunProperties.toString());
		}

		if (cursor.toFirstChild()) {
			do {
				String tagName = cursor.getName().getLocalPart();
				String namespace = cursor.getName().getNamespaceURI();

				if ("run".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'run'");
					makeRun(para, cursor.getObject(), mapRunProperties);

				} else if ("bookmarkStart".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'bookmarkStart'");
					makeBookmarkStart(para, cursor);
				} else if ("bookmarkEnd".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'bookmarkEnd'");
					makeBookmarkEnd(para, cursor);

				} else if ("fn".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'fn'");
					makeFootnote(para, cursor.getObject(), mapRunProperties);

				} else if ("hyperlink".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'hyperlink'");
					makeHyperlink(para, cursor);

				} else if ("image".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'image'");
					makeImage(para, cursor);

				} else if ("object".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'object'");
					makeObject(para, cursor);

				} else if ("page-number-ref".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'page-number-ref'");
					makePageNumberRef(para, cursor);

					// Municode custom...
				} else if ("header-rule".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'header-rule'");
					makeHeaderRule(para, cursor);

				} else if ("footer-rule".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'footer-rule'");
					makeFooterRule(para, cursor);

				} else if ("rule".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'rule'");
					makeRule(para, cursor);

				} else if ("minitoc".equals(tagName)) {
					//log.debug("+ [debug makeParagraph() do... 'minitoc'");
					if (cursor.getTextValue() != null) {
						String instr = cursor.getTextValue();
						buildMiniToc(para, cursor, instr);
					}

				} else if ("p".equals(tagName)) { // handle nested paragraphs (so DocBook-ish)...
					//log.debug("+ [debug makeParagraph() do... 'p' (nested)");
					String htmlstyle = cursor.getAttributeText(DocxConstants.QNAME_HTMLSTYLE_ATT);
					String pagebreak = cursor.getAttributeText(DocxConstants.QNAME_PAGEBREAK_ATT);
					
					if(!StringUtils.isEmpty(cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT))) {
						htmlstyle.concat("; " + cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT));
					}
					
					Map<String, String> mapHtmlStyle = createMapHtmlStyle(htmlstyle, pagebreak);

					// rotatepg...SHOULD THIS GO HERE?
//					if(!StringUtils.isEmpty(mapHtmlStyle.values().toString()) 
//							&& "true".equals(mapHtmlStyle.get("rotatepg"))
//						) {
//						CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
//						CTPageSz pageSz = sectPr.addNewPgSz();
//						pageSz.setH(BigInteger.valueOf(12240)); //12240 Twips = 12240/20 = 612 pt = 612/72 = 8.5"
//						pageSz.setW(BigInteger.valueOf(15840)); //15840 Twips = 15840/20 = 792 pt = 792/72 = 11"
//					}
					
					makeParagraph(para, cursor, mapHtmlStyle);

				} else {
					log.warn("[makeParagraph...para, cursor, mapHtmlStyle]: Unexpected element {" + namespace + "}:"
							+ tagName + " in <p>. Ignored.");
				}
			} while (cursor.toNextSibling());
		}

		cursor.pop();
		return para;
	}

	/**
	 * Construct a page number ("PAGE") complex field.
	 * 
	 * @param para   Paragraph to add the field to
	 * @param cursor
	 */
	/*
	 * private void makePageNumberRef(XWPFParagraph para, XmlCursor cursor) { String
	 * fieldData = "PAGE"; makeSimpleField(para, fieldData); }
	 */

	/**
	 * Makes a simple field within the specified paragraph.
	 * 
	 * @param para      Paragraph to add the field to.
	 * @param fieldData The field data, e.g. "PAGE", "DATE", etc. See 17.16 Fields
	 *                  and Hyperlinks.
	 */
	/*
	 * private void makeSimpleField(XWPFParagraph para, String fieldData) {
	 * CTSimpleField ctField = para.getCTP().addNewFldSimple();
	 * ctField.setInstr(fieldData); }
	 */

	/**
	 * Construct a run within a paragraph.
	 * 
	 * @param para          The output paragraph to add the run to.
	 * @param xml           The <run> element
	 * @param mapProperties Zero or more optional parameters
	 */
	private void makeRun(XWPFParagraph para, XmlObject xml, Map<String, String> mapRunProperties)
			throws DocxGenerationException {
		
		XmlCursor cursor = xml.newCursor();

		String run_fontsize = null;		// general paragraph
		//String run_row_fontsize = null;
		//String run_cell_fontsize = null;

		if (!StringUtils.isEmpty(mapRunProperties.toString())) {

			if (mapRunProperties.containsKey("font-size")) {
				if(StringUtils.isEmpty(mapRunProperties.get("font-size"))) { 
					mapRunProperties.remove("font-size");
				} else {
					run_fontsize = mapRunProperties.get("font-size");
				}
			}
			
			if (mapRunProperties.containsKey("rowFontSize")) {
				if(StringUtils.isEmpty(mapRunProperties.get("rowFontSize"))) { 
					mapRunProperties.remove("rowFontSize");
				} else {
					run_fontsize = mapRunProperties.get("rowFontSize");
				}
			}

			if (mapRunProperties.containsKey("cellFontSize")) {
				if(StringUtils.isEmpty(mapRunProperties.get("cellFontSize"))) { 
					mapRunProperties.remove("cellFontSize");
				} else {
					run_fontsize = mapRunProperties.get("cellFontSize");
				}
			}
		}
		
		/* Eliot's hack... */
		/*
		 * for (String propName : mapRunProperties.keySet()) { String value =
		 * mapRunProperties.get(propName);
		 * 
		 * if (value != null) { // FIXME: This is a quick hack. Need a more general and
		 * elegant way to manage setting of properties. if
		 * (DocxConstants.PROPERTY_PAGEBREAK.equals(propName)) {
		 * if(DocxConstants.PROPERTY_VALUE_CONTINUOUS.equals(value)) {
		 * para.setPageBreak(false); } else { para.setPageBreak(true); } } } }
		 */

		XWPFRun run = para.createRun();
		String styleName = cursor.getAttributeText(DocxConstants.QNAME_STYLE_ATT);
		String styleId = cursor.getAttributeText(DocxConstants.QNAME_STYLEID_ATT);

		if (null != styleName && null == styleId) {
			// Look up the style by name:
			XWPFStyle style = para.getDocument().getStyles().getStyleWithName(styleName);
			if (null != style) {
				styleId = style.getStyleId();
			}
		}

		if (null != styleId) {
			run.setStyle(styleId);
		}

		if(!StringUtils.isEmpty(run_fontsize)) {
			run.setFontSize(Integer.valueOf(run_fontsize));
		}
		
		handleFormattingAttributes(run, xml);

		cursor.toLastAttribute();
		cursor.toNextToken(); // Should be first text or sub-element.
		// In this loop, each different token handler is responsible for positioning
		// the cursor past the thing that was handled such that the only END token
		// is the end for the run element being processed.
		while (TokenType.END != cursor.currentTokenType()) {
			// TokenType tokenType = cursor.currentTokenType(); // For debugging
			if (cursor.isText()) {
				run.setText(cursor.getTextValue());
				cursor.toNextToken();
			} else if (cursor.isAttr()) {
				// Ignore attributes in this context.
			} else if (cursor.isStart()) {
				// Handle element within run
				String name = cursor.getName().getLocalPart();
				String namespace = cursor.getName().getNamespaceURI();

				if ("break".equals(name)) {
					makeBreak(run, cursor);

				} else if ("symbol".equals(name)) {
					makeSymbol(run, cursor);

				} else if ("tab".equals(name)) {
					makeTab(run, cursor);

					// Municode custom...
				} else if ("doDateTime".equals(name)) {
					if (cursor.getTextValue() != null) {
						String instr = cursor.getTextValue();
						makeDateTime(run, cursor, instr);
					} else {
						makeDateTime(run, cursor, "America/New_York;EST");
					}
					cursor.toEndToken(); // Skip this element.
				} else {
					log.error("makeRun(): Unexpected element {" + namespace + "}:" + name + ". Skipping.");
					cursor.toEndToken(); // Skip this element.
				}
				cursor.toNextToken();
			} else if (cursor.isComment() || cursor.isProcinst()) {
				// Silently ignore
				// FIXME: Not sure if we need to do more to skip a comment or processing
				// instruction.
				cursor.toNextToken();
			} else {
				// What else could there be?
				if (cursor.getName() != null) {
					log.error("makeRun(): Unhanded XML token " + cursor.getName().getLocalPart());
				} else {
					log.error("makeRun(): Unhanded XML token " + cursor.currentTokenType());
				}
				cursor.toNextToken();
			}
		}

		cursor.pop();

	}

	// This is an initial "Quick and Dirty" stab to manage <rule/> (not even close
	// to ideal)
	private void makeRule(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
		cursor.push();

		int width = 2; // Default width in inches
		String widthQualifier = "in";
		String widthValQualified = "2in";
		int widthDPI = 0;

		double weight = 0.5; // Default weight in points
		String weightQualifier = "pt";
		String weightValQualified = "0.5pt";
		int weightDPI = 0;

		String widthVal = cursor.getAttributeText(DocxConstants.QNAME_RULE_WIDTH_ATT);
		String widthUnitsVal = cursor.getAttributeText(DocxConstants.QNAME_RULE_WIDTH_UNITS_ATT);
		String weightVal = cursor.getAttributeText(DocxConstants.QNAME_RULE_WEIGHT_ATT);
		String weightUnitsVal = cursor.getAttributeText(DocxConstants.QNAME_RULE_WEIGHT_UNITS_ATT);

		if ((null == widthVal) || (null == widthUnitsVal)) {
			log.info("- [info] No qualified width for rule. Using default of " + width + widthQualifier);
			widthValQualified = (String) widthVal + widthUnitsVal;
		}

		if ((null != widthVal) && (null != widthUnitsVal)) {

			try {
				widthValQualified = (String) widthVal + measure2abbrev(widthUnitsVal);
				widthDPI = (int) Measurement.toPixels(widthValQualified, getDotsPerInch());

			} catch (MeasurementException e) {
				widthDPI = 144; // 2in
				log.error(e.getClass().getSimpleName() + ": " + e.getMessage());
				log.error("Rule (horizontal line) using default width value " + widthDPI + "DPI");
			}
		}

		if ((null == weightVal) || (null == weightUnitsVal)) {
			log.info("- [info] No qualified weight for rule. Using default of " + weight + weightQualifier);
			weightValQualified = (String) weightVal + weightUnitsVal;
		}

		if ((null != weightVal) && (null != weightUnitsVal)) {

			try {
				weightValQualified = (String) weightVal + measure2abbrev(weightUnitsVal);
				weightDPI = (int) Measurement.toPixels(weightValQualified, getDotsPerInch());
			} catch (MeasurementException e) {
				weightDPI = 4; // .05 * 72 = 3.6
				log.error(e.getClass().getSimpleName() + ": " + e.getMessage());
				log.error("Rule (horizontal line) using default weight value " + weightDPI + "DPI");
			}
		}

		XWPFRun run = para.createRun();
		/*
		 * Notes: p->run->pict-> "shape" In the future maybe upgrade this to makeShape()
		 * or makeShapeLine() instead of this hack.
		 * 
		 * Sample code that may be of use: final int EMU = 9525; double width *= EMU;
		 * double height *= EMU; CTInline inline =
		 * run.getCTR().addNewDrawing().addNewInline(); String picXml =
		 * String.format(PICXML, id, blipId, width, height); XmlToken xmlToken = null;
		 * try { xmlToken = XmlToken.Factory.parse(picXml); } catch (XmlException xe) {
		 * LOGGER.error(xe.getMessage(), xe.fillInStackTrace()); } inline.set(xmlToken);
		 * inline.setDistT(0); inline.setDistB(0); inline.setDistL(0);
		 * inline.setDistR(0);
		 */

		// This is the initial "Quick and Dirty" stab at hacking up a 'rule' (not even
		// close to ideal)
		try {
			// @SuppressWarnings("static-access")
			double ruleWidth = Measurement.toInches(widthValQualified, getDotsPerInch());

			run.setFontFamily("Swiss");
			int repeats = (int) (ruleWidth * 14); // about 14 per inch

			String str1 = "_";
			StringBuffer buffer = new StringBuffer(str1);
			for (int i = 0; i < repeats; i++) {
				buffer.append(str1);
			}

			run.setText(buffer.toString());

		} catch (MeasurementException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		cursor.pop();
	}

	/**
	 * date and time: ex. (Created: 2020-04-09 12:54:46 [EST])
	 * 
	 */
	private void makeDateTime(XWPFRun run, XmlCursor cursor, String instr) {
		final String DATE_FORMATTER = "yyyy-MM-dd HH:mm:ss";

		// Examples: instr = "America/New_York;EST"; or simply "America/New_York"
		// NOTE: zoneIdString = ""America/New_York", zoneCustomString = "EST"
		String zoneIdString = "";
		if (instr.endsWith(";")) {
			instr = instr.substring(0, instr.lastIndexOf(';'));
		}

		if (instr.contains(";")) {
			zoneIdString = instr.substring(0, instr.lastIndexOf(';'));
		} else {
			zoneIdString = instr;
		}

		String zoneCustomString = "";
		if (instr.contains(";")) {
			zoneCustomString = instr.substring(instr.lastIndexOf(";") + 1);
		} else {
			zoneCustomString = instr;
		}

		ZoneId zoneId = ZoneId.of(zoneIdString);
		LocalDateTime nowZone = LocalDateTime.now(zoneId);
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern(DATE_FORMATTER);
		String formatDateTime = nowZone.format(formatter);
		String outputDateString = "";

		if (zoneCustomString != "" && zoneCustomString != null) {
			outputDateString = "   Created: " + formatDateTime + " [" + zoneCustomString + "]";
		} else {
			outputDateString = "   Created: " + formatDateTime;
		}

		run.setText(outputDateString);
		run.setFontFamily("Consolas");
		run.setFontSize(6);
	}

	/**
	 * abbreviate measure.
	 * 
	 * @param String "wordy measure qualifier" ex. "inches" or "points"
	 */
	private String measure2abbrev(String wordyMeasure) {
		String measureAbbrev = wordyMeasure;
		switch (measureAbbrev) {
		case "inches":
			measureAbbrev = "in";
			break;
		case "points":
			measureAbbrev = "pt";
			break;
		case "picas":
			measureAbbrev = "pc";
			break;
		default:
			measureAbbrev = "pt";
		}

		return measureAbbrev;
	}

	private void makePageNumberRef(XWPFParagraph para, XmlCursor cursor) {
		// PAGE of NUMPAGES...
		XWPFRun run = para.createRun();
		run.addCarriageReturn();
		para.setAlignment(ParagraphAlignment.CENTER);

		run = para.createRun();
		run.setText("Page ");
		para.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
		run = para.createRun();
		run.setText(" of ");
		para.getCTP().addNewFldSimple().setInstr("NUMPAGES \\* MERGEFORMAT");
	}

	/**
	 * Build a Word TOC at cursor's location
	 * @param para			current WXPFparagraph para
	 * @param cursor		current location in input XML file (SWPX)
	 * @param instr			the Word TOC 'instruction' (ex. TOC &#92;o "1-6" &#92;h &#92;z &#92;u ) 
	 */
	private void buildMiniToc(XWPFParagraph para, XmlCursor cursor, String instr) {
		CTP ctP = para.getCTP();
		CTSimpleField toc = ctP.addNewFldSimple();
		toc.setInstr(instr);
		toc.setDirty(STOnOff.TRUE);
	}

	/**
	 * Clean mapFile entries. - if key = 'htmlstyle' then break it out into
	 * individual entries - trim front and aft removing whitespace, &c.
	 * 
	 * @param mapFile passed in map return mapFile
	 */
	private Map<String, String> cleanupMapEntries(Map<String, String> mapFile) {
		
		if (null != mapFile) {
			for (Map.Entry<String, String> entry : mapFile.entrySet()) {
				
				if ("htmlstyle".equals(entry.getKey())) {
					String sHtmlStyle = entry.getValue();
					String newKey = "";
					String newValue = "";
					
					if (sHtmlStyle.contains(";")) {
						for (String retval : entry.getValue().split(";")) {
							// each retval...
							//System.out.println("...split retval=" + retval);
							Integer iend = -1;
							 newKey = retval.substring(0, retval.indexOf(":"));
							 newValue = retval.substring(retval.indexOf(":") + 1);
							
							if (newValue.contains(";")) {
								iend = newValue.indexOf(";");
								if (iend != -1) {
									newValue = newValue.substring(0, iend);
								}
							}
						}
					} else {
						newKey = sHtmlStyle.substring(0, sHtmlStyle.indexOf(":"));
						newValue = sHtmlStyle.substring(sHtmlStyle.indexOf(":") + 1);
					}
					
					newKey = newKey.trim();
					newValue = newValue.trim();
					mapFile.put(newKey, newValue);
				}
			}
		}
		
		if(mapFile.containsKey("htmlstyle")) {
			mapFile.remove("htmlstyle");
		}
		
		return mapFile;
	}

	/**
	 * Construct a Header rule or horizontal line.
	 * 
	 * @param para   Paragraph to add the field to
	 * @param cursor
	 */
	private void makeHeaderRule(XWPFParagraph para, XmlCursor cursor) {
		para.setBorderTop(Borders.SINGLE);
	}

	/**
	 * Construct a Footer rule or horizontal line.
	 * 
	 * @param para   Paragraph to add the field to
	 * @param cursor
	 */
	private void makeFooterRule(XWPFParagraph para, XmlCursor cursor) {
		para.setBorderBottom(Borders.SINGLE);
	}

	/**
	 * Makes a simple field within the specified paragraph.
	 * 
	 * @param para      Paragraph to add the field to.
	 * @param fieldData The field data, e.g. "PAGE", "DATE", etc. See 17.16 Fields
	 *                  and Hyperlinks.
	 */
	@SuppressWarnings("unused")
	private void makeSimpleField(XWPFParagraph para, String fieldData) {
		CTSimpleField ctField = para.getCTP().addNewFldSimple();
		ctField.setInstr(fieldData);
	}

	/**
	 * Handle Formatting Attributes
	 * 
	 * @param run run to add the field to.
	 * @param xml xml - object
	 */
	private void handleFormattingAttributes(XWPFRun run, XmlObject xml) {
		XmlCursor cursor = xml.newCursor();

		if (cursor.toFirstAttribute()) {
			do {
				String attName = cursor.getName().getLocalPart();
				String attValue = cursor.getTextValue();

				if ("bold".equals(attName)) {
					boolean value = Boolean.parseBoolean(attValue);
					run.setBold(value);
				} else if ("caps".equals(attName)) {
					boolean value = Boolean.parseBoolean(attValue);
					run.setCapitalized(value);
				} else if ("color".equals(attName)) {
					// NOTE: color must be an RGB hex number. May need to translate
					// from color names to RGB values.
					run.setColor(attValue);
				} else if ("double-strikethrough".equals(attName)) {
					boolean value = Boolean.parseBoolean(attValue);
					run.setDoubleStrikethrough(value);
				} else if ("emboss".equals(attName)) {
					boolean value = Boolean.parseBoolean(attValue);
					run.setEmbossed(value);
				} else if ("emphasis-mark".equals(attName)) {
					run.setEmphasisMark(attValue);
				} else if ("expand-collapse".equals(attName)) {
					int percentage = Integer.valueOf(attValue);
					run.setTextScale(percentage);
				} else if ("highlight".equals(attName)) {
					run.setTextHighlightColor(attValue);
				} else if ("imprint".equals(attName)) {
					boolean value = Boolean.parseBoolean(attValue);
					run.setImprinted(value);
				} else if ("italic".equals(attName)) {
					boolean value = Boolean.parseBoolean(attValue);
					run.setItalic(value);
				} else if ("outline".equals(attName)) {
					CTOnOff onOff = CTOnOff.Factory.newInstance();
					onOff.setVal(STOnOff.Enum.forString(attValue));
					run.getCTR().getRPr().setOutline(onOff);
				} else if ("position".equals(attName)) {
					int val = Integer.parseInt(attValue);
					run.setTextPosition(val);
				} else if ("shadow".equals(attName)) {
					boolean value = Boolean.parseBoolean(attValue);
					run.setShadow(value);
				} else if ("small-caps".equals(attName)) {
					boolean value = Boolean.parseBoolean(attValue);
					run.setSmallCaps(value);
				} else if ("strikethrough".equals(attName)) {
					boolean value = Boolean.parseBoolean(attValue);
					run.setStrikeThrough(value);
				} else if ("underline".equals(attName)) {
					UnderlinePatterns value;
					try {
						value = UnderlinePatterns.valueOf(attValue.toUpperCase());
						run.setUnderline(value);
					} catch (Exception e) {
						log.error("- [ERROR] Unrecognized underline value \"" + attValue + "\"");
					}
				} else if ("underline-color".equals(attName)) {
					run.setUnderlineColor(attValue);
				} else if ("underline-theme-color".equals(attName)) {
					run.setUnderlineThemeColor(attValue);
				} else if ("vanish".equals(attName)) {
					boolean value = Boolean.parseBoolean(attValue);
					run.setVanish(value);
				} else if ("vertical-alignment".equals(attName)) {
					run.setVerticalAlignment(attValue);
				}

			} while (cursor.toNextAttribute());
		}

	}

	/**
	 * Make a literal tabl in the run.
	 * 
	 * @param run
	 * @param cursor
	 */
	private void makeTab(XWPFRun run, XmlCursor cursor) {

		run.addTab();
	}

	/**
	 * Make a symbol within a run
	 * 
	 * @param run
	 * @param cursor
	 */
	private void makeSymbol(XWPFRun run, XmlCursor cursor) {
		throw new NotImplementedException("symbol within run not implemented");

	}

	/**
	 * Construct a footnote
	 * 
	 * @param para   the paragraph containing the footnote.
	 * @param cursor Pointing at the &lt;fn> element
	 */
	private void makeFootnote(XWPFParagraph para, XmlObject xml, Map<String, String> mapRunProperties)
			throws DocxGenerationException {

		XmlCursor cursor = xml.newCursor();
		String type = cursor.getAttributeText(DocxConstants.QNAME_TYPE_ATT);

		String htmlstyle = cursor.getAttributeText(DocxConstants.QNAME_HTMLSTYLE_ATT);
		String pagebreak = cursor.getAttributeText(DocxConstants.QNAME_PAGEBREAK_ATT);
		
		if(!StringUtils.isEmpty(cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT))) {
			htmlstyle.concat("; " + cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT));
		}
		
		Map<String, String> mapFNAdditionalParameters = createMapHtmlStyle(htmlstyle, pagebreak);

		XWPFAbstractFootnoteEndnote note = null;
		if ("endnote".equals(type)) {
			note = para.getDocument().createEndnote();
		} else {
			note = para.getDocument().createFootnote();
		}

		// NOTE: The paragraph is not created with any initial paragraph.
		if (cursor.toFirstChild()) {
			do {
				String tagName = cursor.getName().getLocalPart();
				String namespace = cursor.getName().getNamespaceURI();
				if ("p".equals(tagName)) {
					XWPFParagraph p = note.createParagraph();
					makeParagraph(p, cursor, mapFNAdditionalParameters);
				} else if ("table".equals(tagName)) {
					XWPFTable table = note.createTable();
					makeTable(table, cursor.getObject(), mapFNAdditionalParameters);
				} else {
					// There are other body-level things that could go in a footnote but
					// we aren't worrying about them for now.
					log.warn("makeFootnote(): Unexpected element {" + namespace + "}:" + tagName
							+ "' in <fn>. Ignored.");
				}
			} while (cursor.toNextSibling());
		}

		para.addFootnoteReference(note);
		cursor.pop();
	}

	/**
	 * Gets the current ID (i.e., the last one generated).
	 * 
	 * @return Current value of ID counter as a BigInteger.
	 */
	@SuppressWarnings("unused")
	private BigInteger currentId() {
		return new BigInteger(Integer.toString(idCtr));
	}

	/**
	 * Get the next ID for use in result objects.
	 * 
	 * @return Next ID value as a BitInteger
	 */
	private BigInteger nextId() {
		BigInteger id = new BigInteger(Integer.toString(idCtr++));
		return id;
	}

	/**
	 * Make a break within a run.
	 * 
	 * @param run    Run to add the break to
	 * @param cursor Cursor pointing to the &lt;break> element
	 */
	private void makeBreak(XWPFRun run, XmlCursor cursor) throws DocxGenerationException {

		String typeValue = cursor.getAttributeText(DocxConstants.QNAME_TYPE_ATT);
		BreakType type = BreakType.TEXT_WRAPPING;
		if ("line".equals(typeValue) || "textWrapping".equals(typeValue)) {
			// Already set to this
		} else if ("page".equals(typeValue)) {
			type = BreakType.PAGE;
		} else if ("column".equals(typeValue)) {
			type = BreakType.COLUMN;
		} else {
			log.warn("makeBreak(): Unexpected @type value '" + typeValue + "'. Using 'line'.");
		}
		run.addBreak(type);
		// Now move the cursor past the end of the break element
		while (cursor.currentTokenType() != TokenType.END) {
			cursor.toNextToken();
		}
		// At this point, current token is the end of the break element
	}

	/**
	 * Construct a bookmark start
	 * 
	 * @param para
	 * @param cursor
	 */
	private void makeBookmarkStart(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
		CTBookmark bookmark = para.getCTP().addNewBookmarkStart();
		bookmark.setName(cursor.getAttributeText(DocxConstants.QNAME_NAME_ATT));
		BigInteger id = nextId();
		bookmark.setId(id);
		this.bookmarkIdToIdMap.put(cursor.getAttributeText(DocxConstants.QNAME_ID_ATT), id);
	}

	/**
	 * Construct a bookmark end
	 * 
	 * @param doc
	 * @param cursor
	 * @throws DocxGenerationException
	 */
	private void makeBookmarkEnd(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
		CTMarkupRange bookmark = para.getCTP().addNewBookmarkEnd();
		String sourceID = cursor.getAttributeText(DocxConstants.QNAME_ID_ATT);
		BigInteger id = this.bookmarkIdToIdMap.get(sourceID);
		if (id == null) {
			throw new DocxGenerationException("No bookmark start found for bookmark end with ID '" + sourceID + "'");
		} else {
			bookmark.setId(id);
		}
	}

	/**
	 * Construct a hyperlink
	 * 
	 * @param doc
	 * @param cursor
	 */
	private void makeHyperlink(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {

		String href = cursor.getAttributeText(DocxConstants.QNAME_HREF_ATT);

		// Hyperlink's anchor (@w:anchor) points to the name (not ID) of a bookmark.
		//
		// Alternatively, can use the @r:id attribute to point to a relationship
		// element that then points to something, normally an external resource
		// targeted by URI.

		// Convention in simple WP XML is fragment identifiers are to bookmark IDs,
		// while everything else is a URI to an external resource.

		CTHyperlink hyperlink = para.getCTP().addNewHyperlink();
		CTR run = hyperlink.addNewR();
		run.addNewT().setStringValue(cursor.getTextValue());

		// Set the appropriate target:

		if (href.startsWith("#")) {
			// Just a fragment ID, must be to a bookmark
			String bookmarkName = href.substring(1);
			hyperlink.setAnchor(bookmarkName);
		} else {
			// Create a relationship that targets the href and use the
			// relationship's ID on the hyperlink
			// It's not yet clear from the POI API how to create a new relationship for
			// use by an external hyperlink.
			// throw new NotImplementedException("Links to external resources not yet
			// implemented.");
		}

		XWPFHyperlinkRun hyperlinkRun = new XWPFHyperlinkRun(hyperlink, run, para);
		para.addRun(hyperlinkRun);

	}

	/**
	 * Construct an image reference
	 * 
	 * @param doc
	 * @param cursor
	 */
	private void makeImage(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
		cursor.push();

		String imgUrl = cursor.getAttributeText(DocxConstants.QNAME_SRC_ATT);
		if (null == imgUrl) {
			log.error("- [ERROR] No @src attribute for image.");
			return;
		}
		URL url;
		try {
			if (!imgUrl.matches("^\\w+:.+")) {
				String baseUrl = inFile.getParentFile().toURI().toURL().toExternalForm();
				imgUrl = baseUrl + imgUrl;
			}
			url = new URL(imgUrl);
		} catch (MalformedURLException e) {
			log.error("- [ERROR] " + e.getClass().getSimpleName() + " on img/@src value: " + e.getMessage());
			return;
		}
		File imgFile = null;
		try {
			imgFile = new File(url.toURI());
		} catch (URISyntaxException e) {
			// Should never get here.
		}

		String imgFilename = imgFile.getName();
		String imgExtension = FilenameUtils.getExtension(imgFilename).toLowerCase();
		int width = 200; // Default width in pixels
		int height = 200; // Default height in pixels

		int format = getImageFormat(imgExtension);

		if (format == 0) {
			// FIXME: Might be more appropriate to throw an exception here.
			log.error("Unsupported picture: " + imgFilename
					+ ". Expected emf|wmf|pict|jpeg|jpg|png|dib|gif|tiff|eps|bmp|wpg");
			cursor.pop();
			return;
		}

		BufferedImage img = null;
		int intrinsicWidth = 0;
		int intrinsicHeight = 0;
		try {
			// FIXME: Need to limit this to the formats Java2D can read.
			img = ImageIO.read(imgFile);
			intrinsicWidth = img.getWidth();
			intrinsicHeight = img.getHeight();

		} catch (IOException e) {
			log.warn("" + e.getClass().getSimpleName() + " exception loading image file '" + imgFile + "': "
					+ e.getMessage());
		}

		String widthVal = cursor.getAttributeText(DocxConstants.QNAME_WIDTH_ATT);
		String heightVal = cursor.getAttributeText(DocxConstants.QNAME_HEIGHT_ATT);
		boolean goodWidth = false;
		boolean goodHeight = false;

		if ((null != widthVal) && (Measurement.isNumeric(widthVal))) {
			// try {
			// width = (int) Measurement.toPixels(widthVal, getDotsPerInch());
			// (RAYMOND) Turned OFF the conversion from pixels to points in XSLT...
			// thus the width is already in pixels.
			width = Integer.valueOf(widthVal);
			// } catch (MeasurementException e) {
			// log.error(e.getClass().getSimpleName() + ": " + e.getMessage());
			// log.error("Using default width value " + width);
			// width = intrinsicWidth > 0 ? intrinsicWidth : width;
			// }
		} else {
			width = intrinsicWidth > 0 ? intrinsicWidth : width;
		}

		if ((null != heightVal) && (Measurement.isNumeric(heightVal))) {
			// try {
			// height = (int) Measurement.toPixels(heightVal, getDotsPerInch());
			// (RAYMOND) Turned OFF the conversion from pixels to points in XSLT...
			// thus the width is already in pixels.
			height = Integer.valueOf(heightVal);
			// } catch (MeasurementException e) {
			// log.error(e.getClass().getSimpleName() + ": " + e.getMessage());
			// log.error("Using default height value " + height);
			// height = intrinsicHeight > 0 ? intrinsicHeight : height;
			// }
		} else {
			height = intrinsicHeight > 0 ? intrinsicHeight : height;
		}

		// Issue 16: If either dimension is not specified, scale the intrinsic width
		// proportionally.
		if (widthVal == null && heightVal != null && (intrinsicWidth > 0) && goodHeight) {
			double factor = height / intrinsicHeight;
			width = (int) Math.round(intrinsicWidth * factor);
		}
		if (widthVal != null && heightVal == null && (intrinsicHeight > 0) && goodWidth) {
			double factor = (double) width / intrinsicWidth;
			height = (int) Math.round(intrinsicHeight * factor);
		}

		// At this point, the measurement is pixels. If the original specification
		// was also pixels, we need to convert to inches and then back to pixels
		// in order to apply the dots-per-inch value.

		// Word uses a DPI of 72, so if the current dotsPerInch is not 72, we need to
		// adjust the width and height by the difference.

		if (getDotsPerInch() != 72) {
			double factor = 72.0 / getDotsPerInch();
			if (widthVal != null && widthVal.matches("[0-9]+(px)?")) {
				width = (int) Math.round(width * factor);
			}
			if (heightVal != null && heightVal.matches("[0-9]+(px)?")) {
				height = (int) Math.round(height * factor);
			}
		}

		XWPFRun run = para.createRun();

		try {
			run.addPicture(new FileInputStream(imgFile), format, imgFilename, Units.toEMU(width), Units.toEMU(height));
		} catch (Exception e) {
			log.warn("" + e.getClass().getSimpleName() + " exception adding picture for reference '" + imgFile + "': "
					+ e.getMessage());
		}
		cursor.pop();
	}

	/**
	 * Get the current dots-per-inch setting
	 * 
	 * @return Dots (pixels) per inch
	 */
	public int getDotsPerInch() {
		return this.dotsPerInch;
	}

	/**
	 * Set the dots-per-inch to use when converting from pixels to absolute
	 * measurements.
	 * <p>
	 * Typical values are 72 and 96
	 * </p>
	 * 
	 * @param dotsPerInch The dots-per-inch value.
	 */
	public void setDotsPerInch(int dotsPerInch) {
		this.dotsPerInch = dotsPerInch;
	}

	/**
	 * Construct an embedded object.
	 * 
	 * @param para
	 * @param cursor Cursor pointing to an <object> element.
	 */
	private void makeObject(XWPFParagraph para, XmlCursor cursor) throws DocxGenerationException {
		throw new NotImplementedException("Object handling not implemented");
		// cursor.push();
		// cursor.pop();

	}

	/**
	 * Construct an embedded object.
	 * 
	 * @param doc
	 * @param cursor Cursor pointing to an <object> element.
	 */
	private void makeObject(XWPFDocument doc, XmlCursor cursor) throws DocxGenerationException {
		throw new NotImplementedException("Object handling not implemented");
		// cursor.push();
		// cursor.pop();
	}

	
	// NOT USED; USING Docx template table styles instead.
//	private void setTableAlign(XWPFTable table, ParagraphAlignment align) {
//		CTTblPr tblPr = table.getCTTbl().getTblPr(); 
//		CTJc jc = (tblPr.isSetJc() ? tblPr.getJc() : tblPr.addNewJc()); 
//		Enum en = Enum.forInt(align.getValue());
//		jc.setVal(en);
//		
//		log.info("+ [debug] setTableAlign align and en: " + align.toString() + "\t" + en.toString());
//	}

	// Example call to setTableAlign (above)...
	// setTableAlign(table, ParagraphAlignment.CENTER);

	/**
	 * Construct a table.
	 * 
	 * @param table Table object to construct
	 * @param xml   The &lt;table&gt; element
	 * @throws DocxGenerationException
	 */
	private void makeTable(XWPFTable table, XmlObject xml, Map<String, String> mapAdditionalParameters)
			throws DocxGenerationException {
		
		// If the column widths are absolute measurements they can be set on the grid,
		// but if they are proportional, then they have to be set on at least the first
		// row's cells. The table grid is not required (it always reflects the
		// calculated width of the columns, possibly determined by applying percentage
		// table and column widths.
		XmlCursor cursor = xml.newCursor();

		String widthValue = cursor.getAttributeText(DocxConstants.QNAME_WIDTH_ATT);
		if (null != widthValue) {
			table.setWidth(getMeasurementValue(widthValue));
		}

		setTableIndents(table, cursor);

		String styleName = cursor.getAttributeText(DocxConstants.QNAME_STYLE_ATT);
		String styleId = cursor.getAttributeText(DocxConstants.QNAME_STYLEID_ATT);

		if (null != styleName && null == styleId) {
			// Look up the style by name:
			XWPFStyle style = table.getBody().getXWPFDocument().getStyles().getStyleWithName(styleName);
			if (null != style) {
				styleId = style.getStyleId();
			} else {
				// Try to make a style ID out of the style name:
				styleId = styleName.replace(" ", "");
			}
		}
		if (null != styleId) {
			table.setStyleID(styleId);
		}

		TableBorderStyles borderStyles = setTableFrame(table, cursor);

		Map<QName, String> defaults = new HashMap<QName, String>();
		String rowsep = cursor.getAttributeText(DocxConstants.QNAME_ROWSEP_ATT);

		if (rowsep != null) {
			defaults.put(DocxConstants.QNAME_ROWSEP_ATT, rowsep);
		}

		String colsep = cursor.getAttributeText(DocxConstants.QNAME_COLSEP_ATT);

		if (colsep != null) {
			defaults.put(DocxConstants.QNAME_COLSEP_ATT, colsep);
		}

		int borderWidth = 8; // 8 8ths of a point, i.e. 1pt
		int borderSpace = 0; // NOT 8 [Raymond]
		String borderColor = "auto";

		// Rowsep is either 1 or 0
		// Default for new tables is all frames and internal borders so need
		// to explicitly set to none if rowsep or colsep is 0.
		if (rowsep != null || colsep != null) {
			if ("1".equals(rowsep)) {
				borderStyles.setRowSepBorder(borderStyles.getDefaultBorderType());
				table.setInsideHBorder(borderStyles.getRowSepBorder(), borderWidth, borderSpace, borderColor);
			} else {
				borderStyles.setRowSepBorder(XWPFBorderType.NONE);
				table.setInsideHBorder(borderStyles.getRowSepBorder(), 0, 0, borderColor);
			}

			if ("1".equals(colsep)) {
				borderStyles.setColSepBorder(borderStyles.getDefaultBorderType());
				table.setInsideVBorder(borderStyles.getColSepBorder(), borderWidth, borderSpace, borderColor);
			} else {
				borderStyles.setRowSepBorder(XWPFBorderType.NONE);
				table.setInsideVBorder(borderStyles.getRowSepBorder(), 0, 0, borderColor);
			}
		}

		// Not setting a grid on the tables because it only uses absolute
		// measurements. And there's no XWPF API for it.

		// So setting widths on columns, which allows percentages as well as
		// explicit values.
		TableColumnDefinitions colDefs = new TableColumnDefinitions();
		cursor.toChild(DocxConstants.QNAME_COLS_ELEM);
		
		if (cursor.toFirstChild()) {
			do {
				TableColumnDefinition colDef = colDefs.newColumnDef();

				String width = cursor.getAttributeText(DocxConstants.QNAME_COLWIDTH_ATT);
				if (null != width && !width.equals("")) {
					try {
						colDef.setWidth(width, getDotsPerInch());
					} catch (MeasurementException e) {
						log.warn("makeTable(): " + e.getClass().getSimpleName() + " - " + e.getMessage());
					}
				} else {
					colDef.setWidthAuto();
				}
			} while (cursor.toNextSibling());
		}

		// populate the rows and cells.
		cursor = xml.newCursor();

		// Header rows:
		cursor.push();
		if (cursor.toChild(DocxConstants.QNAME_THEAD_ELEM)) {
			if (cursor.toFirstChild()) {
				RowSpanManager rowSpanManager = new RowSpanManager();
				do {
					// Process the rows
					XWPFTableRow row = makeTableRow(table, cursor.getObject(), colDefs, rowSpanManager, defaults);
					row.setRepeatHeader(true);
				} while (cursor.toNextSibling());
			}
		}

		// Body rows:
		cursor = xml.newCursor();

		if (cursor.toChild(DocxConstants.QNAME_TBODY_ELEM)) {
			if (cursor.toFirstChild()) {
				RowSpanManager rowSpanManager = new RowSpanManager();
				do {
					// Process the rows
					XWPFTableRow row = makeTableRow(table, cursor.getObject(), colDefs, rowSpanManager, defaults);

					// Adjust row as needed.
					row.getCtRow(); // For setting low-level properties.
				} while (cursor.toNextSibling());
			}
		}

		table.removeRow(0); // Remove the first row that's always added automatically (FIXME: This may not
							// be needed any more)
	}

	private void setTableIndents(XWPFTable table, XmlCursor cursor) {
		// Should only have left/right or inside/outside values, not both.

		CTTbl ctTbl = table.getCTTbl();
		CTTblPr ctTblPr = (ctTbl.getTblPr());
		if (ctTblPr == null) {
			ctTblPr = ctTbl.addNewTblPr();
		}
		String leftindentValue = cursor.getAttributeText(DocxConstants.QNAME_LEFTINDENT_ATT);
		// There only seems to be a way to set the left indent at the CT* level
//    String rightindentValue = cursor.getAttributeText(DocxConstants.QNAME_RIGHTINDENT_ATT);
//    String insideindentValue = cursor.getAttributeText(DocxConstants.QNAME_LEFTINDENT_ATT);
//    String outsideindentValue = cursor.getAttributeText(DocxConstants.QNAME_RIGHTINDENT_ATT);

		if (leftindentValue != null) {
			CTTblWidth tblWidth = CTTblWidth.Factory.newInstance();
			String value = getMeasurementValue(leftindentValue);
			try {
				tblWidth.setW(new BigInteger(value));
				ctTblPr.setTblInd(tblWidth);
			} catch (Exception e) {
				log.debug("setTableIndents(): leftindentVale \"" + leftindentValue + "\" not an integer", e);
			}
		}

	}

	/**
	 * Get the word measurement value as either a keyword, a percentage, or a twips
	 * integer.
	 * 
	 * @param measurement The measurement to convert
	 * @return Twips value, percentage, or "auto"
	 */
	public String getMeasurementValue(String measurement) {
		String result = "auto";
		if (measurement.endsWith("%") || measurement.equals("auto")) {
			result = measurement;
		} else {
			try {
				long twips = Measurement.toTwips(measurement, getDotsPerInch());
				result = "" + twips;
			} catch (Exception e) {
				log.warn("getMeasurementValue(): " + e.getClass().getSimpleName() + " - " + e.getMessage(), e);
			}
		}
		return result;
	}

	private TableBorderStyles setTableFrame(XWPFTable table, XmlCursor cursor) {
		int frameWidth = 8; // 1pt
		int frameSpace = 0;
		String frameColor = "auto";

		String frameValue = cursor.getAttributeText(DocxConstants.QNAME_FRAME_ATT);

		TableBorderStyles borderStyles = new TableBorderStyles(cursor.getObject());

		XWPFBorderType topBorder = borderStyles.getTopBorder();
		XWPFBorderType bottomBorder = borderStyles.getBottomBorder();
		XWPFBorderType leftBorder = borderStyles.getLeftBorder();
		XWPFBorderType rightBorder = borderStyles.getRightBorder();

		if (frameValue != null) {
			if ("none".equals(frameValue)) {
				topBorder = XWPFBorderType.NONE;
				bottomBorder = XWPFBorderType.NONE;
				leftBorder = XWPFBorderType.NONE;
				rightBorder = XWPFBorderType.NONE;
			} else if ("all".equals(frameValue)) {
				topBorder = getBorderStyle(topBorder, borderStyles.getDefaultBorderType());
				bottomBorder = getBorderStyle(bottomBorder, borderStyles.getDefaultBorderType());
				leftBorder = getBorderStyle(leftBorder, borderStyles.getDefaultBorderType());
				rightBorder = getBorderStyle(rightBorder, borderStyles.getDefaultBorderType());
			} else if ("topbot".equals(frameValue)) {
				topBorder = getBorderStyle(topBorder, borderStyles.getDefaultBorderType());
				bottomBorder = getBorderStyle(bottomBorder, borderStyles.getDefaultBorderType());
				leftBorder = XWPFBorderType.NONE;
				rightBorder = XWPFBorderType.NONE;
			} else if ("sides".equals(frameValue)) {
				topBorder = XWPFBorderType.NONE;
				bottomBorder = XWPFBorderType.NONE;
				leftBorder = getBorderStyle(leftBorder, borderStyles.getDefaultBorderType());
				rightBorder = getBorderStyle(rightBorder, borderStyles.getDefaultBorderType());
			} else if ("top".equals(frameValue)) {
				topBorder = getBorderStyle(topBorder, borderStyles.getDefaultBorderType());
				bottomBorder = XWPFBorderType.NONE;
				leftBorder = XWPFBorderType.NONE;
				rightBorder = XWPFBorderType.NONE;
			} else if ("bottom".equals(frameValue)) {
				topBorder = XWPFBorderType.NONE;
				bottomBorder = getBorderStyle(bottomBorder, borderStyles.getDefaultBorderType());
				leftBorder = XWPFBorderType.NONE;
				rightBorder = XWPFBorderType.NONE;
			}

		}

		if (topBorder != null) {
			table.setTopBorder(topBorder, frameWidth, frameSpace, frameColor);
		}
		if (bottomBorder != null) {
			table.setBottomBorder(bottomBorder, frameWidth, frameSpace, frameColor);
		}
		if (leftBorder != null) {
			table.setLeftBorder(leftBorder, frameWidth, frameSpace, frameColor);
		}
		if (rightBorder != null) {
			table.setRightBorder(rightBorder, frameWidth, frameSpace, frameColor);
		}

		return borderStyles;
	}

	/**
	 * Get the border style, using the default if the explicit style null
	 * 
	 * @param explicitStyle Explicitly-specified border style. May be null
	 * @param defaultType   The default to use if explicit is null
	 * @return The effective border style
	 */
	private XWPFBorderType getBorderStyle(XWPFBorderType explictType, XWPFBorderType defaultType) {
		return (explictType == null ? defaultType : explictType);
	}

	/**
	 * Get the XWPFBorderType for the specified STBorder value.
	 * 
	 * @param borderValue Border value (e.g., "wave").
	 * @return Corresponding XWPFBorderType value or null if there is no
	 *         corresponding value.
	 */
	private XWPFBorderType xwpfBorderType(String borderValue) {

		STBorder.Enum borderStyle = STBorder.Enum.forString(borderValue);

		// There's not a direct correspondence between STBorder int values
		// and XWPFBorderType so just building a switch statement.
		XWPFBorderType xwpfType = null;
		switch (borderStyle.intValue()) {
		case STBorder.INT_DOT_DASH:
			xwpfType = XWPFBorderType.DOT_DASH;
			break;
		case STBorder.INT_DASH_SMALL_GAP:
			xwpfType = XWPFBorderType.DASH_SMALL_GAP;
			break;
		case STBorder.INT_DASH_DOT_STROKED:
			xwpfType = XWPFBorderType.DASH_DOT_STROKED;
			break;
		case STBorder.INT_DASHED:
			xwpfType = XWPFBorderType.DASHED;
			break;
		case STBorder.INT_DOT_DOT_DASH:
			xwpfType = XWPFBorderType.DOT_DOT_DASH;
			break;
		case STBorder.INT_DOTTED:
			xwpfType = XWPFBorderType.DOTTED;
			break;
		case STBorder.INT_DOUBLE:
			xwpfType = XWPFBorderType.DOUBLE;
			break;
		case STBorder.INT_DOUBLE_WAVE:
			xwpfType = XWPFBorderType.DOUBLE_WAVE;
			break;
		case STBorder.INT_INSET:
			xwpfType = XWPFBorderType.INSET;
			break;
		case STBorder.INT_NIL:
			xwpfType = XWPFBorderType.NIL;
			break;
		case STBorder.INT_NONE:
			xwpfType = XWPFBorderType.NONE;
			break;
		case STBorder.INT_OUTSET:
			xwpfType = XWPFBorderType.OUTSET;
			break;
		case STBorder.INT_SINGLE:
			xwpfType = XWPFBorderType.SINGLE;
			break;
		case STBorder.INT_THICK:
			xwpfType = XWPFBorderType.THICK;
			break;
		case STBorder.INT_THICK_THIN_LARGE_GAP:
			xwpfType = XWPFBorderType.THICK_THIN_LARGE_GAP;
			break;
		case STBorder.INT_THICK_THIN_MEDIUM_GAP:
			xwpfType = XWPFBorderType.THICK_THIN_MEDIUM_GAP;
			break;
		case STBorder.INT_THICK_THIN_SMALL_GAP:
			xwpfType = XWPFBorderType.THICK_THIN_SMALL_GAP;
			break;
		case STBorder.INT_THIN_THICK_LARGE_GAP:
			xwpfType = XWPFBorderType.THIN_THICK_LARGE_GAP;
			break;
		case STBorder.INT_THIN_THICK_MEDIUM_GAP:
			xwpfType = XWPFBorderType.THIN_THICK_MEDIUM_GAP;
			break;
		case STBorder.INT_THIN_THICK_SMALL_GAP:
			xwpfType = XWPFBorderType.THIN_THICK_SMALL_GAP;
			break;
		case STBorder.INT_THIN_THICK_THIN_LARGE_GAP:
			xwpfType = XWPFBorderType.THIN_THICK_THIN_LARGE_GAP;
			break;
		case STBorder.INT_THIN_THICK_THIN_MEDIUM_GAP:
			xwpfType = XWPFBorderType.THIN_THICK_THIN_MEDIUM_GAP;
			break;
		case STBorder.INT_THIN_THICK_THIN_SMALL_GAP:
			xwpfType = XWPFBorderType.THIN_THICK_THIN_SMALL_GAP;
			break;
		case STBorder.INT_THREE_D_EMBOSS:
			xwpfType = XWPFBorderType.THREE_D_EMBOSS;
			break;
		case STBorder.INT_THREE_D_ENGRAVE:
			xwpfType = XWPFBorderType.THREE_D_ENGRAVE;
			break;
		case STBorder.INT_TRIPLE:
			xwpfType = XWPFBorderType.TRIPLE;
			break;
		case STBorder.INT_WAVE:
			xwpfType = XWPFBorderType.WAVE;
			break;
		}
		return xwpfType;
	}

	/**
	 * Get the STBorderType.Enum for the specified STBorder value.
	 * 
	 * @param borderValue Border value (e.g., "wave").
	 * @return Corresponding XWPFBorderType value or null if there is no
	 *         corresponding value.
	 */
	private STBorder.Enum stBorderType(XWPFBorderType borderType) {

		// There's not a direct correspondence between STBorder int values
		// and XWPFBorderType so just building a switch statement.
		STBorder.Enum stBorder = null;
		switch (borderType) {
		case DOT_DASH:
			stBorder = STBorder.DOT_DASH;
			break;
		case DASH_SMALL_GAP:
			stBorder = STBorder.DASH_SMALL_GAP;
			break;
		case DASH_DOT_STROKED:
			stBorder = STBorder.DASH_DOT_STROKED;
			break;
		case DASHED:
			stBorder = STBorder.DASHED;
			break;
		case DOT_DOT_DASH:
			stBorder = STBorder.DOT_DOT_DASH;
			break;
		case DOTTED:
			stBorder = STBorder.DOTTED;
			break;
		case DOUBLE:
			stBorder = STBorder.DOUBLE;
			break;
		case DOUBLE_WAVE:
			stBorder = STBorder.DOUBLE_WAVE;
			break;
		case INSET:
			stBorder = STBorder.INSET;
			break;
		case NIL:
			stBorder = STBorder.NIL;
			break;
		case NONE:
			stBorder = STBorder.NONE;
			break;
		case OUTSET:
			stBorder = STBorder.OUTSET;
			break;
		case SINGLE:
			stBorder = STBorder.SINGLE;
			break;
		case THICK:
			stBorder = STBorder.THICK;
			break;
		case THICK_THIN_LARGE_GAP:
			stBorder = STBorder.THICK_THIN_LARGE_GAP;
			break;
		case THICK_THIN_MEDIUM_GAP:
			stBorder = STBorder.THICK_THIN_MEDIUM_GAP;
			break;
		case THICK_THIN_SMALL_GAP:
			stBorder = STBorder.THICK_THIN_SMALL_GAP;
			break;
		case THIN_THICK_LARGE_GAP:
			stBorder = STBorder.THIN_THICK_LARGE_GAP;
			break;
		case THIN_THICK_MEDIUM_GAP:
			stBorder = STBorder.THIN_THICK_MEDIUM_GAP;
			break;
		case THIN_THICK_SMALL_GAP:
			stBorder = STBorder.THIN_THICK_SMALL_GAP;
			break;
		case THIN_THICK_THIN_LARGE_GAP:
			stBorder = STBorder.THIN_THICK_THIN_LARGE_GAP;
			break;
		case THIN_THICK_THIN_MEDIUM_GAP:
			stBorder = STBorder.THIN_THICK_THIN_MEDIUM_GAP;
			break;
		case THIN_THICK_THIN_SMALL_GAP:
			stBorder = STBorder.THIN_THICK_THIN_SMALL_GAP;
			break;
		case THREE_D_EMBOSS:
			stBorder = STBorder.THREE_D_EMBOSS;
			break;
		case THREE_D_ENGRAVE:
			stBorder = STBorder.THREE_D_ENGRAVE;
			break;
		case TRIPLE:
			stBorder = STBorder.TRIPLE;
			break;
		case WAVE:
			stBorder = STBorder.WAVE;
			break;
		}
		return stBorder;
	}

	/**
	 * Construct a table row
	 * 
	 * @param table          The table to add the row to
	 * @param xml            The <row> element to add to the table
	 * @param colDefs        Column definitions
	 * @param rowSpanManager Manages setting vertical spanning across multiple rows.
	 * @param defaults       Defaults inherited from the table (or elsewhere)
	 * @return Constructed row object
	 * @throws DocxGenerationException
	 */
	private XWPFTableRow makeTableRow(XWPFTable table, XmlObject xml, TableColumnDefinitions colDefs,
			RowSpanManager rowSpanManager, Map<QName, String> defaults
			) throws DocxGenerationException {
		
		// NOTE: Future, add the table's map parameters to the call parameters for this method.
		XmlCursor cursor = xml.newCursor();
		XWPFTableRow row = table.createRow();

		String rowFontSize = "";		
		String htmlstyle = cursor.getAttributeText(DocxConstants.QNAME_HTMLSTYLE_ATT);
		
		if(!StringUtils.isEmpty(cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT))) {
			htmlstyle.concat("; " + cursor.getAttributeText(DocxConstants.QNAME_ROTATEPG_ATT));
		}

		Map<String, String> mapRowAdditionalParameters = new HashMap<String, String>();

		if (!StringUtils.isEmpty(htmlstyle)) {
			mapRowAdditionalParameters.put("htmlstyle", htmlstyle);
		}

		mapRowAdditionalParameters = cleanupMapEntries(mapRowAdditionalParameters);
		
		if (mapRowAdditionalParameters.containsKey("font-size")) {
			rowFontSize = String.valueOf(Integer.valueOf(mapRowAdditionalParameters.get("font-size")));
		}
		
		cursor.push();
		cursor.toChild(DocxConstants.QNAME_TD_ELEM);
		int cellCtr = 0;

		do {
			TableColumnDefinition colDef = colDefs.get(cellCtr);
			// Rows always have at least one cell
			// FIXME: At some point the POI API will remove the automatic creation
			// of the first cell in a row.
			XWPFTableCell cell = cellCtr == 0 ? row.getCell(0) : row.addNewTableCell();

			CTTcPr ctTcPr = cell.getCTTc().addNewTcPr();
			String align = cursor.getAttributeText(DocxConstants.QNAME_ALIGN_ATT);
			String rotate = cursor.getAttributeText(DocxConstants.QNAME_ROTATE_ATT);
			String height = cursor.getAttributeText(DocxConstants.QNAME_HEIGHT_ATT);
			String valign = cursor.getAttributeText(DocxConstants.QNAME_VALIGN_ATT);
			String colspan = cursor.getAttributeText(DocxConstants.QNAME_COLSPAN_ATT);
			String rowspan = cursor.getAttributeText(DocxConstants.QNAME_ROWSPAN_ATT);
			String shade = cursor.getAttributeText(DocxConstants.QNAME_SHADE_ATT);

			String cellFontSize = null;
			
			String cell_htmlstyle_att = cursor.getAttributeText(DocxConstants.QNAME_HTMLSTYLE_ATT);

			Map<String, String> mapCellAdditionalParameters = new HashMap<String, String>();
			
			if (!StringUtils.isEmpty(rowFontSize)) {
				cellFontSize = rowFontSize;
			}

			if (!StringUtils.isEmpty(cell_htmlstyle_att)) {

				mapCellAdditionalParameters.put("htmlstyle", cell_htmlstyle_att);
				mapCellAdditionalParameters = cleanupMapEntries(mapCellAdditionalParameters);

				if (!StringUtils.isEmpty(mapCellAdditionalParameters.get("font-size"))) {
					cellFontSize = String.valueOf(Integer.valueOf(mapCellAdditionalParameters.get("font-size")));
				}		
			}

			setCellBorders(cursor, ctTcPr);
			long spanCount = 1; // Default value;

			try {
				String widthValue = cursor.getAttributeText(DocxConstants.QNAME_WIDTH_ATT);
				if (null != widthValue) {
					cell.setWidth(TableColumnDefinition.interpretWidthSpecification(widthValue, getDotsPerInch()));
				} else {
					String width = null;
					width = colDef.getWidth();

					if (colspan != null) {
						try {
							spanCount = Integer.parseInt(colspan);
							// Try to add up the widths of the spanned columns.
							// This is only possible if the values are all percents
							// or are all measurements. Since we don't the actual
							// width of the table itself necessarily, there's no way
							// to reliably convert percentages to explicit widths.
							List<String> spanWidths = new ArrayList<String>();
							boolean allPercents = true;
							boolean allNumbers = true;
							boolean allAuto = true;

							for (int i = cellCtr; i < cellCtr + spanCount; i++) {
								String widthVal = colDefs.get(i).getSpecifiedWidth();
								spanWidths.add(widthVal);
								allPercents = allPercents && widthVal.endsWith("%");
								allNumbers = allNumbers && !widthVal.endsWith("%") && !widthVal.equals("auto");
								allAuto = allAuto && widthVal.equals("auto");
							}

							if (allPercents) {
								double spanPercent = 0;
								for (String cand : spanWidths) {
									String number = cand.substring(0, cand.lastIndexOf("%"));
									try {
										spanPercent += Double.parseDouble(number);
									} catch (NumberFormatException e) {
										log.warn("Calculating width of column-spanning cell: Expected percent value \""
												+ cand + "\" is not numeric.");
									}
								}
								width = "" + spanPercent + "%";
							} else if (allAuto) {
								// Set widths to equal percents so we can calculate span widths.
								int colCount = colDefs.getColumnDefinitions().size();
								double spanPercent = 100.0 / colCount;
								width = "" + spanPercent + "%";
							} else if (allNumbers) {
								int spanMeasurement = 0;
								for (String cand : spanWidths) {
									String number = TableColumnDefinition.interpretWidthSpecification(cand,
											getDotsPerInch());
									try {
										spanMeasurement += Integer.parseInt(number);
									} catch (NumberFormatException e) {
										log.warn("Expected percent value \"" + cand + "\" is not numeric.");
									}
								}
								width = "" + spanMeasurement;
							} else {
								log.warn(
										"Widths of spanned columns are neither all percents or all measurements, cannot calculate exact spanned width");
								log.warn("Widths are \"" + String.join("\", \"", spanWidths) + "\"");
							}

							cell.setWidth(width);
						} catch (Exception e) {
							log.error("makeTableRow(): @colspan value \"" + colspan
									+ "\" is not an integer. Using first column's width.");
						}
					}
					cell.setWidth(width);
					// log.debug("makeTableRow(): Setting width from column definition: " +
					// colDef.getWidth() + " (" + colDef.getSpecifiedWidth() + ")");
				}
			} catch (Exception e) {
				log.error(e.getClass().getSimpleName() + " setting width for column " + (cellCtr + 1) + ": "
						+ e.getMessage(), e);
			}

			if (null != valign) {
				XWPFVertAlign vertAlign = XWPFVertAlign.valueOf(valign.toUpperCase());
				cell.setVerticalAlignment(vertAlign);
			}

			if (null != colspan) {
				try {
					int spanval = Integer.parseInt(colspan);
					CTDecimalNumber spanNumber = CTDecimalNumber.Factory.newInstance();
					spanNumber.setVal(BigInteger.valueOf(spanval));
					// Set the gridspan on the cell to the span count. This will usually
					// set up the width correctly when Word lays out the table
					// regardless of what the nominal column width is. This is because
					// Word infers the table grid from the columns and cells automatically.
					// However, it appears this doesn't always work as expected.
					ctTcPr.setGridSpan(spanNumber);
				} catch (NumberFormatException e) {
					log.warn("Non-numeric value for @colspan: \"" + colspan + "\". Ignored.");
				}
			}

			if (null != rowspan) {
				try {
					int spanval = Integer.parseInt(rowspan);
					CTDecimalNumber spanNumber = CTDecimalNumber.Factory.newInstance();
					spanNumber.setVal(BigInteger.valueOf(spanval));
					rowSpanManager.addColumn(cellCtr, spanval);
					CTVMerge vMerge = CTVMerge.Factory.newInstance();
					vMerge.setVal(STMerge.RESTART);
					ctTcPr.setVMerge(vMerge);
				} catch (NumberFormatException e) {
					log.warn("Non-numeric value for @rowspan: \"" + rowspan + "\". Ignored.");
				}
			}

			if (null != height) {
				String rawHeight = "";
				if (Measurement.isNumeric(height)) {
					rawHeight = height + "px";
				} else {
					rawHeight = height;
				}

				try {
					int heightTwips = (int) Measurement.toTwips(rawHeight, getDotsPerInch());
					row.setHeight(heightTwips);
				} catch (NumberFormatException | MeasurementException e) {
					log.warn("Problem with row @height: \"" + height + "\". Ignored.");
				}
			}

			if (null != shade) {
				try {
					CTShd ctShd = CTShd.Factory.newInstance();
					// <w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>
					ctShd.setFill(shade);
					ctShd.setColor("auto");
					ctShd.setVal(STShd.CLEAR);
					ctTcPr.setShd(ctShd);
				} catch (Exception e) {
					log.warn("Shade value must be a 6-digit hex string, got \"" + shade + "\"");
				}
			}

			if (null != rotate) {
				try {
					// <w:tcPr>
					// <w:tcW w:w="1525" w:type="dxa"/>
					// <w:textDirection w:val="btLr"/>
					// </w:tcPr>

					/*
					 * TextDirection[] textdir = TextDirection.values(); for(int x=0; x<=
					 * textdir.length; x++) { System.out.println("[" + x + "] " + textdir[x]); }
					 */
					// RESULTS: HORIZONTAL, VERTICAL, VERTICAL_270, STACKED

					switch (rotate) {
					case "VERTICAL_270":
						ctTcPr.addNewTextDirection().setVal(STTextDirection.BT_LR);
						break;

					case "VERTICAL":
						ctTcPr.addNewTextDirection().setVal(STTextDirection.LR_TB);
						break;

					case "STACKED":
						ctTcPr.addNewTextDirection().setVal(STTextDirection.LR_TB);
						break;

					case "HORIZONTAL":
						ctTcPr.addNewTextDirection().setVal(STTextDirection.LR_TB);
						break;

					default:
						log.debug("+ [debug] Must fix processing for 'rotate' value:" + rotate);
					}
				} catch (Exception e) {
					log.warn("Bad 'rotate' value " + rotate);
				}
			}

			cursor.push();
			// The first cell of a span will already have a vertical span set for it.
			if (rowspan == null && cursor.toChild(DocxConstants.QNAME_VSPAN_ELEM)) {
				int spansRemaining = rowSpanManager.includeCell(cellCtr);
				if (spansRemaining < 0) {
					log.warn("Found <vspan> when there should not have been one. Ignored.");
				} else {
					ctTcPr.setVMerge(CTVMerge.Factory.newInstance());
				}
			} else {
				if (cursor.toChild(DocxConstants.QNAME_P_ELEM)) {
					do {
						XWPFParagraph p = cell.addParagraph();						
						Map<String, String> mapParaParameters = new HashMap<String, String>();
						
						if (!StringUtils.isEmpty(rowFontSize)) {
							mapParaParameters.put("font-size", rowFontSize);
						}
						if (!StringUtils.isEmpty(cellFontSize)) {
							mapParaParameters.put("font-size", cellFontSize);
						}

						makeParagraph(p, cursor, mapParaParameters);

						if (null != align) {
							if ("JUSTIFY".equalsIgnoreCase(align)) {
								// Issue 18: "BOTH" is the better match to "JUSTIFY"
								align = "BOTH"; // Slight mistmatch between markup and model
							}
							if ("CHAR".equalsIgnoreCase(align)) {
								// I'm not sure this is the best mapping but it seemed close enough
								align = "NUM_TAB"; // Slight mistmatch between markup and model
							}
							ParagraphAlignment alignment = ParagraphAlignment.valueOf(align.toUpperCase());
							p.setAlignment(alignment);
						}
					} while (cursor.toNextSibling());
					// Cells always have at least one paragraph.
					cell.removeParagraph(0);
				}
			}
			cursor.pop();
			cellCtr += spanCount;
		} while (cursor.toNextSibling());

		return row;
	}

//	private static void setRun(XWPFRun run , String fontFamily , int fontSize , String colorRGB , String text , boolean bold , boolean addBreak) {
//		//run.setFontFamily(fontFamily);
//		run.setFontSize(fontSize);
//		//run.setColor(colorRGB);
//		//run.setText(text);
//		//run.setBold(bold);
//		//if (addBreak) run.addBreak();
//	}

	/**
	 * Set the borders on the cells.
	 * 
	 * @param cursor cursor for the table cell markup
	 * @param ctTcPr Table cell style properties
	 * @return
	 */
	private void setCellBorders(XmlCursor cursor, CTTcPr ctTcPr) {

		// log.debug("setCellBorders(): tag is \"" + cursor.getName().getLocalPart() +
		// "\"");
		TableBorderStyles borderStyles = new TableBorderStyles(cursor.getObject());

		if (borderStyles.hasBorders()) {
			CTTcBorders borders = ctTcPr.addNewTcBorders();

			// Borders can be set per edge:

			if (borderStyles.getBottomBorder() != null) {
				CTBorder bottom = borders.addNewBottom();
				STBorder.Enum val = borderStyles.getBottomBorderEnum();
				if (val != null) {
					bottom.setVal(val);
				} else {
					log.warn("setCellBorders(): Failed to get STBorder.Enum value for XWPFBorderStyle \""
							+ borderStyles.getBottomBorder().name() + "\"");
				}
			}
			if (borderStyles.getTopBorder() != null) {
				CTBorder top = borders.addNewTop();
				top.setVal(borderStyles.getTopBorderEnum());
			}
			if (borderStyles.getLeftBorder() != null) {
				CTBorder left = borders.addNewLeft();
				left.setVal(borderStyles.getLeftBorderEnum());
			}
			if (borderStyles.getRightBorder() != null) {
				CTBorder right = borders.addNewRight();
				right.setVal(borderStyles.getRightBorderEnum());
			}
		}
	}

	private void setupStyles(XWPFDocument doc, XWPFDocument templateDoc) throws DocxGenerationException {
		// Load template. For now this is hard coded but will need to be parameterized
		// Copy the template's styles to result document:

		try {
			XWPFStyles newStyles = doc.createStyles();
			newStyles.setStyles(templateDoc.getStyle());
		} catch (IOException e) {
			new DocxGenerationException(e.getClass().getSimpleName() + " reading template DOCX file: " + e.getMessage(),
					e);
		} catch (XmlException e) {
			new DocxGenerationException(
					e.getClass().getSimpleName() + " Copying styles from template doc: " + e.getMessage(), e);
		}
	}

	private void setupNumbering(XWPFDocument doc, XWPFDocument templateDoc) throws DocxGenerationException {
		// Load the template's numbering definitions to the new document

		try {
			XWPFNumbering templateNumbering = templateDoc.getNumbering();
			XWPFNumbering numbering = doc.createNumbering();
			// There is no method to just get all the abstract and concrete
			// numbers or their IDs so we just iterate until we don't get any more

			// Abstract numbers:
			int i = 1;
			XWPFAbstractNum abstractNum = null;
			// Number IDs appear to always be integers starting at 1
			// so we're really just guessing.
			do {
				abstractNum = templateNumbering.getAbstractNum(BigInteger.valueOf(i));
				i++;
				if (abstractNum != null) {
					numbering.addAbstractNum(abstractNum);
				}
			} while (abstractNum != null);

			// Concrete numbers:
			XWPFNum num = null;
			i = 1;
			do {
				num = templateNumbering.getNum(BigInteger.valueOf(i));
				i++;
				if (num != null) {
					numbering.addNum(num);
				}
			} while (num != null);

		} catch (Exception e) {
			new DocxGenerationException(e.getClass().getSimpleName()
					+ " Copying numbering definitions from template doc: " + e.getMessage(), e);
		}

	}

	/**
	 * Set up any custom styles.
	 * 
	 * @param doc Word doc to set up styles for
	 */
	@SuppressWarnings("unused")
	private void setupFootnoteStyles(XWPFDocument doc) throws DocxGenerationException {

		// Styles for footnotes:

		doc.createStyles(); // Make sure we have styles

		CTStyle style = CTStyle.Factory.newInstance();
		style.setStyleId("FootnoteReference");
		style.setType(STStyleType.CHARACTER);
		style.addNewName().setVal("footnote reference");
		style.addNewBasedOn().setVal("DefaultParagraphFont");
		style.addNewUiPriority().setVal(new BigInteger("99"));
		style.addNewSemiHidden();
		style.addNewUnhideWhenUsed();
		style.addNewRPr().addNewVertAlign().setVal(STVerticalAlignRun.SUPERSCRIPT);

		doc.getStyles().addStyle(new XWPFStyle(style));

		style = CTStyle.Factory.newInstance();
		style.setType(STStyleType.PARAGRAPH);
		style.setStyleId("FootnoteText");
		style.addNewName().setVal("footnote text");
		style.addNewBasedOn().setVal("Normal");
		style.addNewLink().setVal("FootnoteTextChar");
		style.addNewUiPriority().setVal(new BigInteger("99"));
		style.addNewSemiHidden();
		style.addNewUnhideWhenUsed();
		CTRPr rpr = style.addNewRPr();
		rpr.addNewSz().setVal(new BigInteger("20"));
		rpr.addNewSzCs().setVal(new BigInteger("20"));

		doc.getStyles().addStyle(new XWPFStyle(style));

		style = CTStyle.Factory.newInstance();
		style.setCustomStyle(STOnOffImpl.X_1);
		style.setStyleId("FootnoteTextChar");
		style.setType(STStyleType.CHARACTER);
		style.addNewName().setVal("Footnote Text Char");
		style.addNewBasedOn().setVal("DefaultParagraphFont");
		style.addNewLink().setVal("FootnoteText");
		style.addNewUiPriority().setVal(new BigInteger("99"));
		style.addNewSemiHidden();
		rpr = style.addNewRPr();
		rpr.addNewSz().setVal(new BigInteger("20"));
		rpr.addNewSzCs().setVal(new BigInteger("20"));

		doc.getStyles().addStyle(new XWPFStyle(style));

	}

	/**
	 * Get the Word-specific format value.
	 * 
	 * @param imgExtension
	 * @return The format or 0 (zero) if the format is not recognized.
	 */
	private int getImageFormat(String imgExtension) {
		int format = 0;

		if ("emf".equals(imgExtension))
			format = XWPFDocument.PICTURE_TYPE_EMF;
		else if ("wmf".equals(imgExtension))
			format = XWPFDocument.PICTURE_TYPE_WMF;
		else if ("pict".equals(imgExtension))
			format = XWPFDocument.PICTURE_TYPE_PICT;
		else if ("jpeg".equals(imgExtension) || "jpg".equals(imgExtension))
			format = XWPFDocument.PICTURE_TYPE_JPEG;
		else if ("png".equals(imgExtension))
			format = XWPFDocument.PICTURE_TYPE_PNG;
		else if ("dib".equals(imgExtension))
			format = XWPFDocument.PICTURE_TYPE_DIB;
		else if ("gif".equals(imgExtension))
			format = XWPFDocument.PICTURE_TYPE_GIF;
		else if ("tiff".equals(imgExtension))
			format = XWPFDocument.PICTURE_TYPE_TIFF;
		else if ("eps".equals(imgExtension))
			format = XWPFDocument.PICTURE_TYPE_EPS;
		else if ("bmp".equals(imgExtension))
			format = XWPFDocument.PICTURE_TYPE_BMP;
		else if ("wpg".equals(imgExtension))
			format = XWPFDocument.PICTURE_TYPE_WPG;

		return format;
	}

	
	/**
	 * Use this after locating the proper sectPr..?
	 */
	/*
	private void changeOrientation(CTSectPr section, String orientation) {
	    CTPageSz pageSize = section.isSetPgSz()? section.getPgSz() : section.addNewPgSz();
	    
	    if (orientation.equals("landscape")) {
	        pageSize.setOrient(STPageOrientation.LANDSCAPE);
	        pageSize.setW(BigInteger.valueOf(842 * 20));
	        pageSize.setH(BigInteger.valueOf(595 * 20));
	    } else {
	        pageSize.setOrient(STPageOrientation.PORTRAIT);
	        pageSize.setH(BigInteger.valueOf(842 * 20));
	        pageSize.setW(BigInteger.valueOf(595 * 20));
	    }
	}
	*/
	
}

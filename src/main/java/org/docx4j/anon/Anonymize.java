package org.docx4j.anon;

import java.util.Map.Entry;

import org.apache.commons.codec.binary.Base64;
import org.docx4j.TraversalUtil;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImageBmpPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImageGifPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImageJpegPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImagePngPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImageTiffPart;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class Anonymize {
	
	private static Logger log = LoggerFactory.getLogger(Anonymize.class);
	
	
	public Anonymize(WordprocessingMLPackage wordMLPackage) {
		
		this.pkg = wordMLPackage;
	}
	
	private WordprocessingMLPackage pkg;
	
	ScrambleText latinizer = null;   
	
	// We'll replace images with 2x2 pixels	
    private static byte[] PNG_IMAGE_DATA;
    private static byte[] GIF_IMAGE_DATA;
    private static byte[] JPEG_IMAGE_DATA;
    private static byte[] BMP_IMAGE_DATA;
    private static byte[] TIF_IMAGE_DATA;
    
    static {
    	
    	PNG_IMAGE_DATA = Base64.decodeBase64("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACAgMAAAAP2OW3AAAADFBMVEUDAP//AAAA/wb//AAD4Tw1AAAACXBIWXMAAAsTAAALEwEAmpwYAAAADElEQVQI12NwYNgAAAF0APHJnpmVAAAAAElFTkSuQmCC");
    	GIF_IMAGE_DATA = Base64.decodeBase64("R0lGODdhAgACAKEEAAMA//8AAAD/Bv/8ACwAAAAAAgACAAACAww0BQA7");
        JPEG_IMAGE_DATA = Base64.decodeBase64(
        						"/9j/4AAQSkZJRgABAQEASABIAAD/4QCMRXhpZgAATU0AKgAAAAgABwEaAAUAAAABAAAAYgEbAAUA"
        						+"AAABAAAAagEoAAMAAAABAAIAAAExAAIAAAASAAAAclEQAAEAAAABAQAAAFERAAQAAAABAAALE1ES"
								+"AAQAAAABAAALEwAAAAAAARlIAAAD6AABGUgAAAPoUGFpbnQuTkVUIHYzLjUuMTAA/9sAQwACAQEC"
								+"AQECAgICAgICAgMFAwMDAwMGBAQDBQcGBwcHBgcHCAkLCQgICggHBwoNCgoLDAwMDAcJDg8NDA4L"
								+"DAwM/9sAQwECAgIDAwMGAwMGDAgHCAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwM"
								+"DAwMDAwMDAwMDAwMDAwM/8AAEQgAAgACAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAAB"
								+"AgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNC"
								+"scEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0"
								+"dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY"
								+"2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//E"
								+"ALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoW"
								+"JDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWG"
								+"h4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp"
								+"6vLz9PX29/j5+v/aAAwDAQACEQMRAD8A/QL4W/sD/ArXfhl4dvr74K/CW8vbzS7ae4uJ/CGnySzy"
								+"NErM7sYiWYkkknkk0UUV5+U/7jR/wR/JHRP4mf/Z");
        BMP_IMAGE_DATA = Base64.decodeBase64("Qk1GAAAAAAAAADYAAAAoAAAAAgAAAAIAAAABABgAAAAAAAAAAAATCwAAEwsAAAAAAAAAAAAABv8AAPz///8AAP//AAMD/w==");
        TIF_IMAGE_DATA = Base64.decodeBase64("");
    }
	
	/*
	 * TODO: everything else, eg OLE embedded objects, WordArt, SmartArt (?)...
	 * VML
	 * EMF, WMF
	 * Detect if present?
	 * 
	 * in which case, return false;  they can still get the docx...
	 * 
	 * List of handled parts .. anything else ..
	 * 
	 */
	public void go() throws Docx4JException {
		
		handleMetadata();
		
//		detectUnsafeParts();
		
		/* content stories:
		 * 
		 * - MDP
		 * - Header/Footer
		 * - Footnotes/Endnotes
		 * - Comments
		 * 
		 * replace with latin text */
		applyScrambleCallbackToParts();
		
		// Next, images
		handleImages();
		
//		detectUnsafeContent();
		
	}

	/**
	 * @throws Docx4JException
	 */
	protected void handleMetadata() throws Docx4JException {
		
		// docProps/app.xml (Extended Properties) 
		if (pkg.getDocPropsExtendedPart()!=null
				&& pkg.getDocPropsExtendedPart().getContents()!=null) {
			
			pkg.getDocPropsExtendedPart().getContents().setCompany(null);
			pkg.getDocPropsExtendedPart().getContents().setManager(null);
			pkg.getDocPropsExtendedPart().getContents().setHeadingPairs(null);
		}
		
		if (pkg.getDocPropsCorePart()!=null
				&& pkg.getDocPropsCorePart().getContents()!=null) {
			
			pkg.getDocPropsCorePart().getContents().setCategory(null);
			pkg.getDocPropsCorePart().getContents().setCreator(null);
			pkg.getDocPropsCorePart().getContents().setDescription(null);
			pkg.getDocPropsCorePart().getContents().setIdentifier(null);
			pkg.getDocPropsCorePart().getContents().setKeywords(null);
			pkg.getDocPropsCorePart().getContents().setSubject(null);
			pkg.getDocPropsCorePart().getContents().setTitle(null);
		}

		if (pkg.getDocPropsCustomPart()!=null) {
			// Remove this
			pkg.getRelationshipsPart().removePart(pkg.getDocPropsCustomPart().getPartName());
		}
	}
	
    /**
     * This method replaces images with 2x2 pixels (which Word scales appropriately)
     * 
     * @throws InvalidFormatException
     */
    private void handleImages() 
    		throws InvalidFormatException {
        
	    // Apply map to headers/footers
		for (Entry<PartName, Part> entry : pkg.getParts().getParts().entrySet()) {

			Part p = entry.getValue(); 

			if (p instanceof ImagePngPart
					|| p instanceof ImageGifPart
					|| p instanceof ImageJpegPart
					|| p instanceof ImageBmpPart
					|| p instanceof ImageTiffPart
					
					// TODO: eps
					
					) {
				
				((BinaryPart)p).setBinaryData(PNG_IMAGE_DATA);
				
			} 
			
		}
    }	
	
    private void applyScrambleCallbackToParts() 
    		throws InvalidFormatException {

		latinizer = new ScrambleText(pkg);
    	
    	
	    // Apply map to MDP                
		toLatin( pkg.getMainDocumentPart() );        							
        
	    // Apply map to headers/footers
		for (Entry<PartName, Part> entry : pkg.getParts().getParts().entrySet()) {

			Part p = entry.getValue(); 

			if (p instanceof HeaderPart) {
	    		toLatin( (HeaderPart)p );        							
			}

			if (p instanceof FooterPart) {
	    		toLatin( (FooterPart)p );        							
			}
			
		}
        
	    // endnotes/footnotes
		if (pkg.getMainDocumentPart().getFootnotesPart()!=null) {
			toLatin( pkg.getMainDocumentPart().getFootnotesPart() );
		}
		if (pkg.getMainDocumentPart().getEndNotesPart()!=null) {
			toLatin( pkg.getMainDocumentPart().getEndNotesPart() );
		}
		
		
        // Comments
		if (pkg.getMainDocumentPart().getCommentsPart()!=null) {
			toLatin( pkg.getMainDocumentPart().getCommentsPart() );
		}
   		
		return;

    }  	

	public void toLatin(JaxbXmlPart p) {
		
		new TraversalUtil(p.getJaxbElement(), latinizer);
		
	}
    
	
	

	
}

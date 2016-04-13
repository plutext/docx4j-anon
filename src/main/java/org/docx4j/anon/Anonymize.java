package org.docx4j.anon;

import java.io.StringWriter;
import java.util.List;
import java.util.Map.Entry;
import java.util.Random;

import org.apache.commons.codec.binary.Base64;
import org.docx4j.Docx4J;
import org.docx4j.TextUtils;
import org.docx4j.TraversalUtil;
import org.docx4j.TraversalUtil.CallbackImpl;
import org.docx4j.fonts.RunFontSelector;
import org.docx4j.fonts.RunFontSelector.RunFontCharacterVisitor;
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
import org.docx4j.wml.P;
import org.docx4j.wml.Text;
import org.w3c.dom.Document;

import com.thedeanda.lorem.Lorem;
import com.thedeanda.lorem.LoremIpsum;


public class Anonymize {
	
	public Anonymize(WordprocessingMLPackage wordMLPackage) {
		
		this.pkg = wordMLPackage;
	}
	
	private WordprocessingMLPackage pkg;
	
	private static Lorem lorem = LoremIpsum.getInstance();
	Latinizer latinizer = null;   
	
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
	 * docProps/app.xml (Extended Properties) : leave as is
	 * 
	 * docProps/core.xml: remove company info
	 * 
	 * 
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
	public void go() throws InvalidFormatException {

		/* content stories:
		 * 
		 * - MDP
		 * - Header/Footer
		 * - Footnotes/Endnotes
		 * - Comments
		 * 
		 * replace with latin text */
		applyLatinCallbackToParts();
		
		// Next, images
		handleImages();
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
	
    private void applyLatinCallbackToParts() 
    		throws InvalidFormatException {

		latinizer = new Latinizer();   
    	
    	
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
    
    public static class Latinizer extends CallbackImpl {
    	
    	
    	String latinText = null;
    	int beginIndex;
    	
    	Random random = new Random();
    	
    	@Override
		public List<Object> apply(Object o) {
    		
			
			if (o instanceof P) {
				
				P p = (P)o;
				
				StringWriter out = new StringWriter();
				try {
					TextUtils.extractText(p, out);
				} catch (Exception e) {
					e.printStackTrace();
				}
				
				String s = out.toString();
				int slenRqd = s.length();
				
				StringBuffer replacement = new StringBuffer();
				int len = 0;
				
				do
				{
					// A bit of effort to get enough text
					
					int wordsNeeded = Math.round((slenRqd-len)/8) + 1; // always at least one word!
					String latin = lorem.getWords(wordsNeeded,wordsNeeded);
					len += latin.length();
					replacement.append(latin);
					
//					System.out.println(len + ", " + slenRqd);
										
				} while (len < slenRqd);
				
				latinText = replacement.toString();
				beginIndex = 0;
				
			}

			if (o instanceof Text) {
				
				Text t = (Text)o;
				int tLen = t.getValue().length();
				
//				t.setValue(latinText.substring(beginIndex, beginIndex+tLen));
				t.setValue(unicodeRangeToFont(t.getValue(), latinText));
				
//				System.out.println(t.getValue());
				
				beginIndex += tLen;
			}
			
			return null;
		}
    	
    	
    	private char getRandom(char rangeLower, char rangeUpper) {
    		
    		    		
	    	char result = (char)(rangeLower + random.nextInt((int)rangeUpper-(int)rangeLower));
	    	
    		return result;
    		
    	}
    	
        private String unicodeRangeToFont(String text, String latinText) {
        	
		    vis.createNew();
        	        	    	
        	if (text==null) {
        		return null; 
        	}
        	for (int i = 0; i < text.length(); i=text.offsetByCodePoints(i, 1)){
        		
        	    char c = text.charAt(i);
        	    
        	    System.out.println(Integer.toHexString(c));
        	    
        	    if (Character.isHighSurrogate(c)) {

        		    System.out.println("high");    		    
    				vis.addCodePointToCurrent(text.codePointAt(i));
    				        	    	
        	    }
        	    else if (c==' ' ) {
        	    	
        	    	// Add it to existing
        	    	vis.addCharacterToCurrent(c);
        	    	
        	    } else {
        		    
//        		    System.out.println(c);    		    
        		    
        		    /* .. Basic Latin
        		     * 
        		     * http://webapp.docx4java.org/OnlineDemo/ecma376/WordML/rFonts.html says 
        		     * @ascii (or @asciiTheme) is used to format all characters in the ASCII range 
        		     * (0 - 127)
        		     */
            	    if (c>='\u0041' && c<='\u00FA') // A-Z 
            	    {
            	    	vis.addCharacterToCurrent( latinText.substring(i, i+1).charAt(0));
            	    	
            	    } else if (c>='\u0061' && c<='\u007A') // a-z 
                	    {
                	    	vis.addCharacterToCurrent( latinText.substring(i, i+1).charAt(0));
                	    	
            	    } else if (c>='\u0000' && c<='\u007F') 
                	    {
                	    	vis.addCharacterToCurrent( c );
                	    	
            	    } else 
            	    	
		        		    // ..  Latin-1 Supplement
		            	    if (c>='\u00A0' && c<='\u00FF') 
		            	    {
		            	    	/* hAnsi (or hAnsiTheme if defined), with the following exceptions:
		        					If hint is eastAsia, the following characters use eastAsia (or eastAsiaTheme if defined): A1, A4, A7 – A8, AA, AD, AF, B0 – B4, B6 – BA, BC – BF, D7, F7
		        					If hint is eastAsia and the language of the run is either Chinese Traditional or Chinese Simplified, the following characters use eastAsia (or eastAsiaTheme if defined): E0 – E1, E8 – EA, EC – ED, F2 – F3, F9 – FA, FC
		        					*/
		            	    	
		            	    	vis.addCharacterToCurrent(c);
		            	    	
		            	    } else 
		        		    // ..  Latin Extended-A, Latin Extended-B, IPA Extensions
		            	    if (c>='\u0100' && c<='\u02AF') 
		            	    {
		            	    	/* hAnsi (or hAnsiTheme if defined), with the following exception:
		        					If hint is eastAsia, and the language of the run is either Chinese Traditional or Chinese Simplified, 
		        					or the character set of the eastAsia (or eastAsiaTheme if defined) font is Chinese5 or GB2312 
		        					then eastAsia (or eastAsiaTheme if defined) font is used.
		        					*/
		            	    	
		            	    	
		            	    	/*
		            	    	 * 
		            	    	 * 
		            	    	 */
		            	    	//c = getRandom('\u0100','\u02AF');
		            	    	
		            	    	vis.addCharacterToCurrent(c);
		            	    	
		            	    } else 
		            	    if (c>='\u02B0' && c<='\u04FF') 
		            	    {
		            	    	
		            	    	c = getRandom('\u02B0' , '\u04FF');
		            	    	vis.addCharacterToCurrent(c);
		            	    	
		            	    }
		            	    else if (c>='\u0590' && c<='\u07BF') 
		            	    {
		            	    	// Arabic: not sure how well this works.  Maybe it does with the right fonts installed?
		            	    	
		            	    	//System.out.println("Arabic");
		            	    	
			            	    c = getRandom('\u0590' , '\u07BF'); 
		            	    	vis.addCharacterToCurrent(c);
		            	    	
		            	    }
		            	    else if (c>='\u1100' && c<='\u11FF') 
		            	    {
			            	    c = getRandom('\u1100' , '\u11FF');
		            	    	vis.addCharacterToCurrent(c);
		            	    	
		            	    } else if (c>='\u1E00' && c<='\u1EFF') 
		            	    {
			            	    c = getRandom('\u1E00' , '\u1EFF');
		            	    	vis.addCharacterToCurrent(c);
		            	    }
		            	    else if (c>='\u2000' && c<='\u2EFF') 
		            	    {
			            	    //c = getRandom('\u2000' , '\u2EFF');
		            	    	vis.addCharacterToCurrent(c);
		            	    	
		            	    }
		            	    else if (c>='\u2F00' && c<='\uDFFF') 
		            	    {
		            	    	// Japanese is in here
		            	    	
			            	    c = getRandom('\u2F00' , '\uDFFF');
		            	    	vis.addCharacterToCurrent(c);
		            	    }
		            	    else if (c>='\uE000' && c<='\uF8FF') 
		            	    {
		            	    	
			            	    c = getRandom('\uE000' , '\uF8FF');
		            	    	vis.addCharacterToCurrent(c);
		            	    }
		            	    else if (c>='\uF900' && c<='\uFAFF') 
		            	    {
			            	    c = getRandom('\uF900' , '\uFAFF');
		            	    	vis.addCharacterToCurrent(c);
		            	    } else 
		        		    // ..  Alphabetic Presentation Forms
		            	    if (c>='\uFB00' && c<='\uFB4F') 
		            	    {
			            	    c = getRandom('\uFB00' , '\uFB4F');
		            	    	vis.addCharacterToCurrent(c);
		            	    } else if (c>='\uFB50' && c<='\uFDFF') {
		            	    	
			            	    c = getRandom('\uFB50' , '\uFDFF');
		            	    	vis.addCharacterToCurrent(c);
		            	    	
		            	    } else if (c>='\uFE30' && c<='\uFE6F') {
		            	    	
			            	    c = getRandom('\uFE30' , '\uFE6F');
		            	    	vis.addCharacterToCurrent(c);
		            	    	
		            	    } else if (c>='\uFE70' && c<='\uFEFE') {
		
			            	    c = getRandom('\uFE70' , '\uFEFE');
		            	    	vis.addCharacterToCurrent(c);
		
		            	    } else if (c>='\uFF00' && c<='\uFFEF') {
		
			            	    c = getRandom('\uFF00' , '\uFFEF');
		            	    	vis.addCharacterToCurrent(c);
		            	    	
		            	    } else {
		        				
		            	    	vis.addCharacterToCurrent(c);
		            	    	
		            	    	
            	    }
            	    
            	    
        	    }
        	} 
        	
        	return (String)vis.getResult();
        }
    	
    	RunFontCharacterVisitor vis = new RunFontCharacterVisitor() {
			
			StringBuilder sb = new StringBuilder(1024); 
			
			@Override
			public void setDocument(Document document) {}
			

			public void addCharacterToCurrent(char c) {
		    	sb.append(c);		
			}
			
			@Override
			public void addCodePointToCurrent(int cp) {
				sb.append(
						new String(Character.toChars(cp)));
			}


			@Override
			public Object getResult() {
				return sb.toString();
			}

			private RunFontSelector runFontSelector;
			@Override
			public void setRunFontSelector(RunFontSelector runFontSelector) {
				this.runFontSelector = runFontSelector;
			}

			@Override
			public void setFallbackFont(String fontname) {}


			@Override
			public void finishPrevious() {
				
			}


			@Override
			public void createNew() {
				sb = new StringBuilder(1024); 
			}


			@Override
			public void setMustCreateNewFlag(boolean val) {
				// TODO Auto-generated method stub
				
			}


			@Override
			public boolean isReusable() {
				// TODO Auto-generated method stub
				return false;
			}


			@Override
			public void fontAction(String fontname) {
				// TODO Auto-generated method stub
				
			}
    	};
        
	}
	
	

	public static void main(String[] args) throws Docx4JException {

//        String inputfilepath = System.getProperty("user.dir") + "/UN-Declaration.docx";
        String inputfilepath = System.getProperty("user.dir") + "/sample-docx.docx";
		
        String outputfilepath = System.getProperty("user.dir") + "/OUT_Anon.docx";
        
        WordprocessingMLPackage pkg = Docx4J.load(new java.io.File(inputfilepath));	
        
        Anonymize anon = new Anonymize(pkg);
        anon.go();
        
        Docx4J.save(pkg, new java.io.File(outputfilepath));
	}
	
}

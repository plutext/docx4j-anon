import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.util.Set;
import java.util.Map.Entry;

import org.apache.commons.io.FileUtils;
import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.anon.Anonymize;
import org.docx4j.anon.AnonymizeResult;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;


public class CorpusAnon {

	private final static String DIR_IN = System.getProperty("user.dir") + "/corpus/";

	private final static String DIR_OUT = System.getProperty("user.dir") + "/OUT/";
	
	private final static String DIR_HANDLED = System.getProperty("user.dir") + "/corpus-handled/";
	
	private static final String DIR_OK = "ok";
	private static final String DIR_LEAKS = "leaks";
	private static final String DIR_ERRORS = "errors";
	
	private int oks = 0;
	private int leaks = 0;
	private int errors = 0;
	
	private StringBuffer sbLeaks = new StringBuffer();
	
	public static void main(String[] args) throws Exception {

		CorpusAnon corpusAnon = new CorpusAnon();
		
		corpusAnon.createDirs();
		corpusAnon.walk(DIR_IN);
		
		System.out.println(corpusAnon.sbLeaks.toString());
		
		System.out.println("leaks: " + corpusAnon.leaks);
		System.out.println("errors: " + corpusAnon.errors);
		System.out.println("oks: " + corpusAnon.oks);
	}
		
	private void createDirs() throws IOException {
		
		// create OK, leak dirs
		FileUtils.forceMkdir(new File(DIR_OUT+DIR_OK));
		FileUtils.forceMkdir(new File(DIR_OUT+DIR_LEAKS));
		FileUtils.forceMkdir(new File(DIR_OUT+DIR_ERRORS));
		
	}

	int docNum = 1;
	
    public void walk( String path ) throws IOException {

        File root = new File( path );
        File[] list = root.listFiles();

        if (list == null) return;

        for ( File f : list ) {
            if ( f.isDirectory() ) {
                walk( f.getAbsolutePath() );
                //System.out.println( "Dir:" + f.getAbsoluteFile() );
            }
            else {
//                System.out.println( "File:" + f.getAbsoluteFile() );
                
                if (f.getName().endsWith("docx")
                		|| f.getName().endsWith("docm")) {
                	
                	try {
						handle(f) ;

						FileUtils.moveFile(f, new File(DIR_HANDLED  + f.getName()));

						
					} catch (Docx4JException e) {

						
						if (e.getMessage().startsWith("This file seems to be a binary doc")) {
							FileUtils.copyFile(f, new File(DIR_OUT+DIR_ERRORS+"/" + f.getName()+".doc"));
							
							// rename the original
							FileUtils.moveFile(f, new File(f.getAbsolutePath()+".doc"));
							
						} else {
							errors++;

							e.printStackTrace();
							FileUtils.copyFile(f, new File(DIR_OUT+DIR_ERRORS+"/" + f.getName()));
							
							File file = new File(DIR_OUT+DIR_ERRORS+"/" + f.getName() + "err.txt");
							PrintStream ps = new PrintStream(file);
							e.printStackTrace(ps);
							ps.close();
						}
					} 
                	
                	docNum++;
                }
            }
        }
    }	
	
    
	private void handle(File fIn) throws Docx4JException {

		System.out.println("\n\n " + docNum + " Processing " + fIn.getName() + "\n\n");

		WordprocessingMLPackage pkg = null;
		try {
			pkg = Docx4J.load(fIn);			
		} catch (ClassCastException e) {
			// eg dodgy docx: CustomXmlDataStoragePart cannot be cast to org.docx4j.openpackaging.parts.CustomXmlDataStoragePropertiesPart
			throw new Docx4JException(e.getMessage(), e);
		}

		Anonymize anon = new Anonymize(pkg);
		AnonymizeResult result = anon.go();
		

		
		String lang = "null";
		if (result.getLang()!=null
				&& result.getLang().getVal()!=null
				&& result.getLang().getVal().length()>=2) {
			
			lang = result.getLang().getVal().substring(0, 2); // we want the language only
			
		}
		
		if (result.isOK()) {
			
			oks++;
			
			System.out.println("document successfully anonymised.");
			
			File dir = new File(DIR_OUT+DIR_OK+"/"+lang);
			dir.mkdirs();
			
			Docx4J.save(pkg, new java.io.File(DIR_OUT+DIR_OK+"/"+lang + "/"+ fIn.getName()));
			
		} else {
			
			leaks++;

			// Report
			reportLeak("\n\n REPORT for " + fIn.getName() + "\n\n");
			
			File dir = new File(DIR_OUT+DIR_LEAKS+"/"+lang);
			dir.mkdirs();
			
			String outputfilepath = DIR_OUT+DIR_LEAKS+"/"+lang + "/"+ fIn.getName();
			
			Docx4J.save(pkg, new java.io.File(outputfilepath));
			
			reportLeak("document partially anonymised; please check " + outputfilepath);
			
			if (result.getUnsafeParts().size()>0) {
				reportLeak("The following parts may leak info:");
				for(Part p :  result.getUnsafeParts()) {
					reportLeak(p.getPartName().getName() + ", of type " + p.getClass().getName() );
				}
			}
			
			// unsafe objects
			if (result.hasAnyUnsafeObjects()) {
				reportLeak("The following objects may leak info:");
				for(Entry<Part, Set<Object>> entry :  result.getUnsafeObjectsByPart().entrySet()) {
					
					Part p = entry.getKey();
					
					if ( !entry.getValue().isEmpty()) {
						reportLeak(p.getPartName().getName() + ", of type " + p.getClass().getName() );
						
						for (Object o : entry.getValue() ) {
							
							if (o instanceof String ) {
								reportLeak((String)o);
							} else if (o instanceof org.docx4j.math.CTOMathPara) { 
								reportLeak("math");						
							} else {
								reportLeak(o.getClass().getName());
								try {
									reportLeak(XmlUtils.marshaltoString(o));							
								} catch (Exception e) {
									reportLeak(e.getMessage());
									e.printStackTrace();
								}
							}
							
						}
					}
					
				}
			}
			System.out.println("\n\n .. end REPORT for " + fIn.getName()  + "\n\n");
			
		}
		
		if (result.getFieldsPresent().size()>0) {
			
			for (String s : result.getFieldsPresent()) {
				System.out.println(s);
			}
		}

		
	}
	
	private void reportLeak(String message) {
		
		System.out.println(message);
		sbLeaks.append(message + "\n");
		
	}
    
	
}

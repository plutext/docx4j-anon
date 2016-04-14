package org.docx4j.anon;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class Example {

	

public static void main(String[] args) throws Docx4JException {

//  String inputfilepath = System.getProperty("user.dir") + "/UN-Declaration.docx";
  String inputfilepath = System.getProperty("user.dir") + "/sample-docx.docx";
	
  String outputfilepath = System.getProperty("user.dir") + "/OUT_Anon.docx";
  
  WordprocessingMLPackage pkg = Docx4J.load(new java.io.File(inputfilepath));	
  
  Anonymize anon = new Anonymize(pkg);
  anon.go();
  
  Docx4J.save(pkg, new java.io.File(outputfilepath));
}

}

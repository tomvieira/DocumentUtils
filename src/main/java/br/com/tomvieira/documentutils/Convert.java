package br.com.tomvieira.documentutils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 *
 * @author Wellington
 */
public class Convert {

    public static void htmlToDocx() throws Exception {
        
        InputStream xhtml = new FileInputStream("xhtml/relatorio.xhtml");

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                .createPackage();
        XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);
        wordMLPackage.getMainDocumentPart().getContent()
                .addAll(XHTMLImporter.convert(xhtml, null));

        System.out.println(XmlUtils.marshaltoString(wordMLPackage
                .getMainDocumentPart().getJaxbElement(), true, true));
        FileOutputStream out = new FileOutputStream("xhtml/out.docx");
        wordMLPackage.save(out);
    }
}

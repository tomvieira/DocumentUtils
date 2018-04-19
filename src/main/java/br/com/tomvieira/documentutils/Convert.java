package br.com.tomvieira.documentutils;

import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 *
 * @author Wellington
 */
public class Convert {

    public static void htmlToDocx() throws Exception {

        String xhtml = "<div>"
                + "<p>The <b>quick</b> <span style=\"font-size: 14pt;\">brown</span> fox...</p>"
                + "<p>Paragraph 2</p>" + "</div>";

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                .createPackage();
        XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);
        wordMLPackage.getMainDocumentPart().getContent()
                .addAll(XHTMLImporter.convert(xhtml, null));

        System.out.println(XmlUtils.marshaltoString(wordMLPackage
                .getMainDocumentPart().getJaxbElement(), true, true));

    }
}

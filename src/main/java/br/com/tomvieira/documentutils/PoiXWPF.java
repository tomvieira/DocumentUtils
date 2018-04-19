package br.com.tomvieira.documentutils;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

/**
 *
 * @author Wellington
 */
public class PoiXWPF {

    public static void merge(List<InputStream> documents, OutputStream dest) throws Exception {
        XWPFDocument documentBase = new XWPFDocument();
        CTBody bodyBase = documentBase.getDocument().getBody();
        for (InputStream document : documents) {
            OPCPackage srcPackage = OPCPackage.open(document);
            XWPFDocument srcDocument = new XWPFDocument(srcPackage);
            CTBody srcBody = srcDocument.getDocument().getBody();
            appendBody(bodyBase, srcBody);
        }
        documentBase.write(dest);
    }

    private static void appendBody(CTBody src, CTBody append) throws Exception {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        String appendString = append.xmlText(optionsOuter);
        String srcString = src.xmlText();
        String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
        String mainPart = srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));
        String sufix = srcString.substring(srcString.lastIndexOf("<"));
        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
        CTBody makeBody = CTBody.Factory.parse(prefix + mainPart + addPart + sufix);
        src.set(makeBody);
    }
}

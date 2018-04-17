package br.com.tomvieira.documentutils;

import com.oracle.webservices.internal.api.message.ContentType;
import com.sun.xml.internal.messaging.saaj.util.ByteInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import static org.docx4j.fonts.FontUtils.target;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.org.apache.poi.util.IOUtils;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTAltChunk;

/**
 *
 * @author Wellington
 */
public class DocumentUtil {

    private static final String CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    public OutputStream joinDocuments(List<InputStream> documents) {
        try {            
            WordprocessingMLPackage target = WordprocessingMLPackage.load(documents.get(0));
            for (InputStream document : documents) {
                insertDocx(target.getMainDocumentPart(), IOUtils.toByteArray(document));
            }
            SaveToZipFile saver = new SaveToZipFile(target);
            saver.save(os);
        } catch (Exception ex) {
            Logger.getLogger(DocumentUtil.class.getName()).log(Level.SEVERE, null, ex);
        }

        return null;
    }

    public void mergeDocx(InputStream s1, InputStream s2, OutputStream os) throws Exception {
        insertDocx(target.getMainDocumentPart(), IOUtils.toByteArray(s2));
    }

    private static void insertDocx(MainDocumentPart main, byte[] bytes) throws Exception {
        AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(new PartName("/part" + (chunk++) + ".docx"));
        afiPart.setContentType(new ContentType(CONTENT_TYPE));
        afiPart.setBinaryData(bytes);
        Relationship altChunkRel = main.addTargetPart(afiPart);

        CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();
        chunk.setId(altChunkRel.getId());

        main.addObject(chunk);
    }

}

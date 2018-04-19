package br.com.tomvieira.documentutils;

import java.io.File;
import java.io.InputStream;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.logging.Level;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTAltChunk;
import java.util.logging.Logger;
import org.apache.commons.io.IOUtils;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Parts;
import org.docx4j.openpackaging.parts.WordprocessingML.AltChunkType;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart.AddPartBehaviour;
import org.docx4j.wml.ContentAccessor;

/**
 *
 * @author Wellington
 */
public class DocumentUtil {

    private static final String CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private static long chunk = 0;

    public void joinDocuments(List<InputStream> documents) {
        try {
            WordprocessingMLPackage target = criarDocumentoVazio();             
            
            for (InputStream document : documents) {
                insertDocx(target, target.getMainDocumentPart(), IOUtils.toByteArray(document));
                IOUtils.closeQuietly(document);
            }
            exportarArquivo(target, "juntado.docx");
        } catch (Exception ex) {
            Logger.getLogger(DocumentUtil.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

//    public void mergeDocx(InputStream s1, InputStream s2, OutputStream os) throws Exception {
//        WordprocessingMLPackage target = lerDocumento(s1);
//        insertDocx(target.getMainDocumentPart(), IOUtils.toByteArray(s2));
//    }
    private static void insertDocx(WordprocessingMLPackage doc, MainDocumentPart main, byte[] bytes) throws Exception {        
        main.addAltChunk(AltChunkType.WordprocessingML, bytes);        
        //main.convertAltChunks();
    }

    
//    public AlternativeFormatInputPart addAltChunk(AltChunkType type, InputStream is) throws Docx4JException {
//
//        AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(type);
//        Relationship altChunkRel = this.addTargetPart(afiPart, AddPartBehaviour.RENAME_IF_NAME_EXISTS);
//        // now that its attached to the package ..
//        afiPart.registerInContentTypeManager();
//
//        afiPart.setBinaryData(is);
//
//        // .. the bit in document body 
//        CTAltChunk ac = Context.getWmlObjectFactory().createCTAltChunk();
//        ac.setId(altChunkRel.getId());
//        if (this instanceof ContentAccessor) {
//            ((ContentAccessor) this).getContent().add(ac);
//        } else {
//            throw new Docx4JException(this.getClass().getName() + " doesn't implement ContentAccessor");
//        }
//
//        return afiPart;
//    }

    public WordprocessingMLPackage lerDocumento(InputStream file) {
        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                    .load(file);
            return wordMLPackage;
        } catch (Docx4JException ex) {
            Logger.getLogger(DocumentUtil.class.getName()).log(Level.SEVERE, null, ex);
        }
        return null;
    }

    public WordprocessingMLPackage criarDocumentoVazio() {
        try {
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
            return wordPackage;
        } catch (InvalidFormatException ex) {
            Logger.getLogger(DocumentUtil.class.getName()).log(Level.SEVERE, null, ex);
        }
        return null;
    }

    public MainDocumentPart getCorpoDocumento(WordprocessingMLPackage wordPackage) {
        MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();
        return mainDocumentPart;
    }

    public void exportarArquivo(WordprocessingMLPackage wordPackage, String nomeArquivo) {
        try {
            //System.out.println(wordPackage.getMainDocumentPart().getXML());
            File exportFile = new File(nomeArquivo);
//            wordPackage.save(exportFile);
            SaveToZipFile saver = new SaveToZipFile(wordPackage);
            saver.save(exportFile);
        } catch (Docx4JException ex) {
            Logger.getLogger(DocumentUtil.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

}

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
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

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
                insertDocx(target,target.getMainDocumentPart(), IOUtils.toByteArray(document));
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

    private static void insertDocx(WordprocessingMLPackage doc,MainDocumentPart main, byte[] bytes) throws Exception {
        
        Parts docParts = doc.getParts();
		Set<PartName> docPartsNames = docParts.getParts().keySet();
		Set<String> plainPartNames = new HashSet<String>();
		for (PartName pn : docPartsNames) {
			plainPartNames.add(pn.getName());
		}

		String partName = null;
		int index = 0;
		do {
			partName = "/part" + index + ".docx";
			index++;
		} while (plainPartNames.contains(partName));

		/*
		 * Now add the bottom file as another part to the top package, and add a
		 * CTAltChunk to the main document of the top package that references
		 * this new part.
		 */

		AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(
				new PartName(partName));
        
        
        
        //AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(new PartName("/part" + (chunk++) + ".docx"));
        afiPart.setContentType(new ContentType(CONTENT_TYPE));
        afiPart.setBinaryData(bytes);
        System.out.println(bytes.toString());
        Relationship altChunkRel = main.addTargetPart(afiPart);

        CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();
        chunk.setId(altChunkRel.getId());

        main.addObject(chunk);
        //main.convertAltChunks();
    }

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

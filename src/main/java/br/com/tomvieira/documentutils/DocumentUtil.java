package br.com.tomvieira.documentutils;

import java.io.File;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

/**
 *
 * @author Wellington
 */
public class DocumentUtil {

    public WordprocessingMLPackage lerDocumento(File file) {
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
            File exportFile = new File(nomeArquivo);
            wordPackage.save(exportFile);
        } catch (Docx4JException ex) {
            Logger.getLogger(DocumentUtil.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

}


import br.com.tomvieira.documentutils.DocumentUtil;
import br.com.tomvieira.documentutils.DocxUtil;
import br.com.tomvieira.documentutils.PoiXWPF;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.Test;

/**
 *
 * @author tom
 */
public class Teste {

//    @Test
//    public void criandoArquivoTeste() {
//        DocumentUtil documentUtil = new DocumentUtil();
//        WordprocessingMLPackage documento = documentUtil.criarDocumentoVazio();
//        documentUtil.getCorpoDocumento(documento).addParagraphOfText("Este é um texto de teste");
//        documentUtil.exportarArquivo(documento, "final2.docx");
//    }
    @Test
    public void juntaArquivos() {
        try {
            DocumentUtil documentUtil = new DocumentUtil();
            List<InputStream> documents = new ArrayList<>();            
            documents.add(new FileInputStream("texto_introdutorio.docx"));
            documents.add(new FileInputStream("capa.docx"));
            documentUtil.joinDocuments(documents);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Teste.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    //@Test
    public void juntaArquivos2() {
        try {
            WordprocessingMLPackage topPackage = WordprocessingMLPackage.load(new File("final.docx"));
            DocxUtil.merge(topPackage, new File("final2.docx"), new File("juntado.docx"));            
        } catch (Exception ex) {
            Logger.getLogger(Teste.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    //@Test
    public void juntarComPoi(){
        try {
            FileOutputStream out = new FileOutputStream(new File("juntado.docx"));
            PoiXWPF.merge(new FileInputStream("capa.docx"),new FileInputStream("texto_introdutorio.docx") , out);
            out.flush();
            
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Teste.class.getName()).log(Level.SEVERE, null, ex);
        } catch (Exception ex) {
            Logger.getLogger(Teste.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}

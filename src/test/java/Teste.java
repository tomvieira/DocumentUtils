
import br.com.tomvieira.documentutils.DocumentUtil;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.bind.JAXBException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.Test;

/**
 *
 * @author tom
 */
public class Teste {

    @Test
    public void criandoArquivoTeste() {
        DocumentUtil documentUtil = new DocumentUtil();
        WordprocessingMLPackage documento = documentUtil.criarDocumentoVazio();
        documentUtil.getCorpoDocumento(documento).addParagraphOfText("Este é um texto de teste");
        documentUtil.exportarArquivo(documento, "final.docx");
    }
}

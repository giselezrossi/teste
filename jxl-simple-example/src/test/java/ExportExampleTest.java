import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.junit.Test;


public class ExportExampleTest {

	@Test
    public void testImportacao(){
        try {
            OutputStream output = new FileOutputStream("c:\\temp\\planilha.xls");
            Exporter exporter = new Exporter();
            exporter.export(output);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        
    }
}

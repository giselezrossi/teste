import java.io.IOException;
import java.io.OutputStream;
import java.util.Locale;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class Exporter {
    
    public void export(OutputStream output){
        
        WorkbookSettings sets = new WorkbookSettings();
        sets.setLocale(new Locale("pt", "BR"));
        sets.setEncoding("ISO-8859-1");
        sets.setFormulaAdjust(true);
        sets.setDrawingsDisabled(true);
        
        try {
            WritableWorkbook book = Workbook.createWorkbook(output, sets);
            WritableSheet planilha = book.createSheet("teste", 0);
            planilha.getSettings().setProtected(true);

            createHeader(planilha);
            createBody(planilha);
            
            book.write();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    	

    }
    
    private void createHeader(WritableSheet planilha) {
        WritableFont fontCell = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
        WritableCellFormat alignmentCelula = new WritableCellFormat(fontCell);
        try {
            alignmentCelula.setAlignment(Alignment.CENTRE);
            alignmentCelula.setBackground(Colour.GREY_25_PERCENT);
        } catch (WriteException e) {
            System.out.println(e.getMessage());
        }
        
        createColumn(planilha, 0, 0, "Name", 4000, alignmentCelula);      
        createColumn(planilha, 1, 0, "Age", 4000, alignmentCelula);
    }
    
    private void createBody(WritableSheet planilha){
        WritableFont fontCell = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
        WritableCellFormat alignmentCelula = new WritableCellFormat(fontCell);
        
        try {
            alignmentCelula.setAlignment(Alignment.CENTRE);
        } catch (WriteException e) {
            System.out.println(e.getMessage());
        }
       
        createColumn(planilha, 0, 1, "gisele", 4000, alignmentCelula);
        createColumn(planilha, 1, 1, "24", 4000, alignmentCelula);
    }
    
    private void createColumn(WritableSheet planilha, int column, int row, String value, int size, WritableCellFormat font) {
        
        Label celulaTexto = new Label(column, row, value, font);
        try {
            planilha.addCell(celulaTexto);
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }

        CellView view = planilha.getColumnView(column);
        view.setSize(size);
        planilha.setColumnView(column, view);
        
    }

}

package convertpdftoexcel;

import com.aspose.pdf.Document;
import com.aspose.pdf.ExcelSaveOptions;
import fileManager.Selector;
import java.io.File;
import javax.swing.JOptionPane;

public class ConvertPDFToExcel {

    public static void main(String[] args) {
        try {
            //Pega Arquivo PDF
            JOptionPane.showMessageDialog(null, "Escolha a seguir o arquivo em PDF a ser convertido:");
            File file = Selector.selectFile("", "PDF", ".pdf");

            if (file != null && file.exists()) {
                //Converte para Excel
                File excelFile = convertPdfToExcel(file);
                JOptionPane.showMessageDialog(null, "O arquivo foi salvo em: \n" + excelFile.getAbsolutePath());
            } else {
                throw new Error("Você deve escolher um arquivo PDF válido!");
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "Ocorreu um erro Java: " + e, "Erro!", JOptionPane.ERROR_MESSAGE);
        } catch (Error e) {
            JOptionPane.showMessageDialog(null, "Erro: " + e.getMessage(), "Erro!", JOptionPane.ERROR_MESSAGE);
        }
        
        System.exit(0);
    }

    public static File convertPdfToExcel(File pdfFile) {
        Document doc = new Document(pdfFile.getAbsolutePath());
        // Set Excel options
        ExcelSaveOptions options = new ExcelSaveOptions();
        // Set output format
        options.setFormat(ExcelSaveOptions.ExcelFormat.XLSX);
        // Set minimizing option
        options.setMinimizeTheNumberOfWorksheets(true);

        //Setr new file xlsx
        File newFile = new File(pdfFile.getAbsolutePath().replaceAll(".PDF", ".pdf").replaceAll(".pdf", ".xlsx"));

        // Convert PDF to XLSX
        doc.save(newFile.getAbsolutePath(), options);

        return newFile;
    }

}

import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        String inputExcelPath = "data.xlsx";
        String templateExcelPath = "template/Report Pass Masuk GAC.xlsx";
        String outputExcelPath = "Report_Pass_Masuk_GAC_Output.xlsx";

        System.out.println("Starting Excel processing...");

        UploadProcessor processor = new UploadProcessor();
        try {
            processor.processExcel(inputExcelPath, templateExcelPath, outputExcelPath);
            System.out.println("Program finished successfully.");
            System.out.println("Generated report: " + outputExcelPath);
        } catch (IOException e) {
            System.err.println("An error occurred during Excel processing: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
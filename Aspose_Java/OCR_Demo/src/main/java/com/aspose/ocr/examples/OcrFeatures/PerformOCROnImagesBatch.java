package com.aspose.ocr.examples.OcrFeatures;
import com.aspose.ocr.*;
import com.aspose.ocr.examples.Utils;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;

public class PerformOCROnImagesBatch {
    public static void main(String[] args) {
        //SetLicense.main(null);
        // ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(PerformOCROnImagesBatch.class);
        
        try {
            // Method 1: Using resource URL (recommended for files in JAR)
            URL resourceUrl = PerformOCROnImagesBatch.class
                    .getClassLoader()
                    .getResource("OCR/photo_1.jpg");
            
            if (resourceUrl == null) {
                System.err.println("Resource 'OCR/photo_1.jpg' not found in classpath");
                return;
            }
            
            String imagePath1 = resourceUrl.getPath();
            
            // Alternative Method 2: Copy resource to temp file if path doesn't work
            /*
            InputStream resourceStream = PerformOCROnImagesBatch.class
                    .getClassLoader()
                    .getResourceAsStream("OCR/p.png");
            
            if (resourceStream == null) {
                System.err.println("Resource 'OCR/p.png' not found in classpath");
                return;
            }
            
            // Create temp file
            File tempFile = File.createTempFile("ocr_temp", ".png");
            tempFile.deleteOnExit();
            
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                byte[] buffer = new byte[1024];
                int bytesRead;
                while ((bytesRead = resourceStream.read(buffer)) != -1) {
                    fos.write(buffer, 0, bytesRead);
                }
            }
            
            String imagePath1 = tempFile.getAbsolutePath();
            */
            
            // Create api instance
            AsposeOCR api = new AsposeOCR();
            
            // Set preprocessing filters
            PreprocessingFilter filters = new PreprocessingFilter();
            filters.add(PreprocessingFilter.AutoSkew());
            
            // Create OcrInput object with correct InputType for image
            OcrInput input = new OcrInput(InputType.SingleImage, filters);  // Changed back to SingleImage
            input.add(imagePath1);
            
            System.out.println("Processing file: " + imagePath1);
            
            // Recognize images/PDF
            ArrayList<RecognitionResult> results = api.Recognize(input);
            
            if (results.isEmpty()) {
                System.out.println("No results returned from OCR processing");
            } else {
                for (RecognitionResult result : results) {
                    System.out.println("---------------------------------");
                    System.out.println("Result: " + result.recognitionText);
                }
            }
            
        } catch (IOException e) {
            System.err.println("IOException occurred: " + e.getMessage());
            e.printStackTrace();
        } catch (Exception e) {
            System.err.println("Unexpected error: " + e.getMessage());
            e.printStackTrace();
        }
        // ExEnd:1
    }
}

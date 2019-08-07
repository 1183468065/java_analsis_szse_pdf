package factory;

import org.apache.pdfbox.text.PDFTextStripper;

import java.io.IOException;

public class PDFReaderFactory {

    private static PDFTextStripper textStripper;

    private PDFReaderFactory() {
    }

    public static PDFTextStripper getPDFTextStripper() {
        if (textStripper == null) {
            try {
                textStripper = new PDFTextStripper();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return textStripper;
    }
}

package com.emailSender;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import com.lowagie.text.Document;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Phrase;
import com.lowagie.text.pdf.ColumnText;
import com.lowagie.text.pdf.PdfContentByte;
import com.lowagie.text.pdf.PdfPageEventHelper;
import com.lowagie.text.pdf.PdfWriter;

public class AllJavaFilesToPdf {

    public static void main(String[] args) {
        String baseDir = "D:/studyPurpose/JavaExamples/src";  // 📌 Update your src path
        String outputPdfPath = "AllJavaFilesOrganized.pdf";

        try {
            // Get all .java files
            List<Path> javaFiles = Files.walk(Paths.get(baseDir))
                    .filter(path -> path.toString().endsWith(".java"))
                    .toList();

            // Group files by package (folder name)
            Map<String, List<Path>> groupedByPackage = new TreeMap<>();
            for (Path file : javaFiles) {
                String relativePath = Paths.get(baseDir).relativize(file).toString();
                String packageName = relativePath.contains(File.separator)
                        ? relativePath.substring(0, relativePath.lastIndexOf(File.separator))
                        : "(default)";
                groupedByPackage.computeIfAbsent(packageName, k -> new ArrayList<>()).add(file);
            }

            // Create PDF document
            Document document = new Document(PageSize.A4, 36, 36, 54, 36);
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(outputPdfPath));

            // Add page numbers
            writer.setPageEvent(new PdfPageNumberHelper());
            document.open();

            Font packageFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 16, Color.DARK_GRAY);
            Font fileFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12, Color.BLUE);
            Font codeFont = FontFactory.getFont(FontFactory.COURIER, 9, Color.BLACK);

            for (String pkg : groupedByPackage.keySet()) {
                // Section: Package name
                Paragraph packageTitle = new Paragraph("Package: " + pkg.replace("\\", "."), packageFont);
                packageTitle.setSpacingBefore(10f);
                packageTitle.setSpacingAfter(10f);
                document.add(packageTitle);

                for (Path file : groupedByPackage.get(pkg)) {
                    document.add(new Paragraph("File: " + file.getFileName(), fileFont));
                    document.add(new Paragraph(" ")); // spacing

                    List<String> lines;
                    try {
                        lines = Files.readAllLines(file, java.nio.charset.StandardCharsets.UTF_8);
                    } catch (java.nio.charset.MalformedInputException e) {
                        lines = Files.readAllLines(file, java.nio.charset.StandardCharsets.ISO_8859_1);
                    }

                    for (String line : lines) {
                        document.add(new Paragraph(line, codeFont));
                    }

                    document.newPage();  // page break after each file
                }
            }

            document.close();
            System.out.println("✅ Organized PDF with package, files, and page numbers created: " + outputPdfPath);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // Helper class to add page numbers
    static class PdfPageNumberHelper extends PdfPageEventHelper {
        Font pageNumberFont = FontFactory.getFont(FontFactory.HELVETICA, 8, Font.ITALIC, Color.GRAY);

        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            PdfContentByte cb = writer.getDirectContent();
            Phrase footer = new Phrase("Page " + writer.getPageNumber(), pageNumberFont);
            ColumnText.showTextAligned(cb, Element.ALIGN_CENTER,
                    footer,
                    (document.right() - document.left()) / 2 + document.leftMargin(),
                    document.bottom() - 10, 0);
        }
    }
}

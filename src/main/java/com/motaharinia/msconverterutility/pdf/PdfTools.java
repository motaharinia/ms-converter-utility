package com.motaharinia.msconverterutility.pdf;


import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.motaharinia.msconverterutility.pdf.dto.CsvDto;
import com.motaharinia.msconverterutility.pdf.utils.ExcelToHtmlConverter;
import io.woo.htmltopdf.HtmlToPdf;
import io.woo.htmltopdf.HtmlToPdfObject;
import io.woo.htmltopdf.PdfOrientation;
import io.woo.htmltopdf.PdfPageSize;
import org.apache.commons.io.FileUtils;
import org.jetbrains.annotations.NotNull;

import java.awt.*;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;

/**
 * @author https://github.com/motaharinia<br>
 * کلاس ابزارهای مربوط به pdf
 */

//https://kb.itextpdf.com/home/it7kb/ebooks/itext-7-converting-html-to-pdf-with-pdfhtml/chapter-7-frequently-asked-questions-about-pdfhtml/how-to-convert-html-containing-arabic-hebrew-characters-to-pdf
//https://github.com/wooio/htmltopdf-java
public interface PdfTools {


    /**
     * متد تولید pdf از رشته html
     *
     * @param html رشته html
     * @return خروجی: خروجی بایت pdf
     * @throws IOException خطا
     */
    @NotNull
    static ByteArrayOutputStream generateFromHtml(@NotNull String html, @NotNull PdfPageSize pdfPageSize, @NotNull PdfOrientation pdfOrientation) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        HtmlToPdf.create()
                .object(HtmlToPdfObject.forHtml(html))
                .pageSize(pdfPageSize)
                .orientation(pdfOrientation)
                .convert().transferTo(byteArrayOutputStream);
        return byteArrayOutputStream;
    }


    /**
     * csv از رشته pdf متد تولید
     *
     * @param csvDto csv مدل تنظیمات
     * @throws IOException خطا
     */
    static ByteArrayOutputStream generateFromCsv(@NotNull CsvDto csvDto) throws IOException, DocumentException {

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        int columnSize = 0;
        String[] splittedData;

        //گرفتن محتوای فایل به صورت استرینگ
        String data = new String(csvDto.getByteArrayInputStream().readAllBytes(), StandardCharsets.UTF_8);

        //ایجاد سند جدید
        Document document = new Document();
        PdfWriter.getInstance(document, byteArrayOutputStream);
        // landscape تنظیم سند به صورت
        document.setPageSize(csvDto.getPageSize());
        document.setMargins(0, 0, 20, 20);
        document.open();

        //تنظیم فونت
        Font headerFont = new Font(BaseFont.createFont(csvDto.getCsvHeaderStyleDto().getFontPath(), BaseFont.IDENTITY_H, BaseFont.EMBEDDED), csvDto.getCsvHeaderStyleDto().getFontSize(), csvDto.getCsvHeaderStyleDto().getFontStyle(), csvDto.getCsvHeaderStyleDto().getFontColor());
        Font bodyFont = new Font(BaseFont.createFont(csvDto.getCsvBodyStyleDto().getFontPath(), BaseFont.IDENTITY_H, BaseFont.EMBEDDED), csvDto.getCsvBodyStyleDto().getFontSize(), csvDto.getCsvBodyStyleDto().getFontStyle(), csvDto.getCsvBodyStyleDto().getFontColor());

        //بررسی csv
        if (!data.isEmpty()) {
            splittedData = data.split(System.lineSeparator());
            columnSize = splittedData[0].split(csvDto.getSeparator()).length;
        } else
            throw new RuntimeException("csv data is empty");

        PdfPTable table = new PdfPTable(columnSize);
        table.setRunDirection(csvDto.getDocumentRightToLeft() ? PdfWriter.RUN_DIRECTION_RTL : PdfWriter.RUN_DIRECTION_LTR);

        //افزودن هدر و بادی
        for (int i = 0; i < splittedData.length; i++) {
            String[] rowCells = splittedData[i].split(csvDto.getSeparator());
            for (int j = 0; j < columnSize; j++) {
                Paragraph paragraph = new Paragraph(rowCells[j]);
                PdfPCell c1 = new PdfPCell();
                c1.setPaddingBottom(20);
                if (i == 0 && csvDto.getHasHeader()) {
                    paragraph.setFont(headerFont);
                    paragraph.setAlignment(csvDto.getCsvHeaderStyleDto().getAlignment());
                    c1.setHorizontalAlignment(csvDto.getCsvHeaderStyleDto().getAlignment());
                    c1.addElement(paragraph);
                    c1.setBorderColor(csvDto.getCsvHeaderStyleDto().getBorderColor());
                    c1.setBackgroundColor(csvDto.getCsvHeaderStyleDto().getBackgroundColor());
                    table.addCell(c1);
                    if (splittedData.length > 1)
                        table.setHeaderRows(1);
                    continue;
                }
                paragraph.setFont(bodyFont);
                c1.addElement(paragraph);
                c1.setBorderColor(csvDto.getCsvBodyStyleDto().getBorderColor());
                c1.setHorizontalAlignment(csvDto.getCsvBodyStyleDto().getAlignment());
                c1.setBackgroundColor(csvDto.getCsvBodyStyleDto().getBackgroundColor());
                table.addCell(c1);
            }

        }

        document.add(table);
        document.close();
        return byteArrayOutputStream;
    }

    private static String getFont(String name) {
        String fonts[] = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();

        for (int i = 0; i < fonts.length; i++) {
            if (name.equals(fonts[i]))
                return FileUtils.getFile(fonts[i]).getAbsolutePath().concat(".ttf");
        }
        throw new RuntimeException("font not found");
    }


    /**
     * excel از pdf متد تولید
     *
     * @param byteArrayInputStream excel بایت ورودی
     * @throws IOException خطا
     */
    static ByteArrayOutputStream generateFromExcel(@NotNull ByteArrayInputStream byteArrayInputStream, @NotNull PdfPageSize pdfPageSize, @NotNull PdfOrientation pdfOrientation) throws IOException {

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        PrintWriter out = new PrintWriter(byteArrayOutputStream);

        ExcelToHtmlConverter toHtml = ExcelToHtmlConverter.create(byteArrayInputStream, out);
        toHtml.setCompleteHTML(true);
        toHtml.printPage();

        byteArrayInputStream.close();

        //Pdf تولید شده به Html تبدیل
        return generateFromHtml(new String(byteArrayOutputStream.toByteArray(), StandardCharsets.UTF_8), pdfPageSize, pdfOrientation);
    }
}

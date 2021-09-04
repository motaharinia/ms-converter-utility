package com.motaharinia.msconverterutility.pdf;



import io.woo.htmltopdf.HtmlToPdf;
import io.woo.htmltopdf.HtmlToPdfObject;
import org.jetbrains.annotations.NotNull;

import java.io.*;

/**
 * @author https://github.com/motaharinia<br>
 * کلاس ابزارهای مربوط به pdf
 */

//https://kb.itextpdf.com/home/it7kb/ebooks/itext-7-converting-html-to-pdf-with-pdfhtml/chapter-7-frequently-asked-questions-about-pdfhtml/how-to-convert-html-containing-arabic-hebrew-characters-to-pdf
//https://github.com/wooio/htmltopdf-java
public interface PdfTools {


    /**
     * متد تولید pdf از رشته html
     * @param html رشته html
     * @return خروجی: خروجی بایت pdf
     * @throws IOException خطا
     */
    @NotNull
    static ByteArrayOutputStream generateFromHtml(@NotNull String html) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        HtmlToPdf.create()
                .object(HtmlToPdfObject.forHtml(html))
                .convert().transferTo(byteArrayOutputStream);
        return byteArrayOutputStream;
    }
}

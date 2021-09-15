package com.motaharinia.msconverterutility.pdf;


import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Element;
import com.itextpdf.text.PageSize;
import com.motaharinia.msconverterutility.pdf.dto.CsvDto;
import com.motaharinia.msconverterutility.pdf.dto.CsvStyleDto;
import io.woo.htmltopdf.PdfOrientation;
import io.woo.htmltopdf.PdfPageSize;
import org.apache.commons.io.FileUtils;
import org.junit.jupiter.api.*;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.Locale;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.fail;

/**
 * @author https://github.com/motaharinia<br>
 * کلاس تست PdfTools
 */
@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
class PdfToolsUnitTest {


    /**
     * این متد مقادیر پیش فرض قبل از هر تست این کلاس تست را مقداردهی اولیه میکند
     */
    @BeforeEach
    void beforeEach() {
        Locale.setDefault(new Locale("fa", "IR"));
    }

    /**
     * این متد بعد از هر تست این کلاس اجرا میشود
     */
    @AfterEach
    void afterEach() {
        Locale.setDefault(Locale.US);
    }

    @Order(1)
    @Test
    void generateFromHtmlTest() {
        try {

            ClassLoader classLoader = getClass().getClassLoader();
            File file = new File(classLoader.getResource("static/pdf/testhtml/test.html").getFile());
            String html = Files.readString(file.toPath(), StandardCharsets.UTF_8);

            try (OutputStream outputStream = new FileOutputStream("test_converted/PdfToolsUnitTest_generateFromHtmlTest.pdf")) {
                PdfTools.generateFromHtml(html, PdfPageSize.A4, PdfOrientation.PORTRAIT).writeTo(outputStream);
            }

        } catch (Exception ex) {
            fail(ex.toString());
        }
    }

    @Order(2)
    @Test
    void generateFromCsvTest() {
        try {

            ClassLoader classLoader = getClass().getClassLoader();
            File file = new File(classLoader.getResource("static/pdf/testCsv/test.csv").getPath());

            // csv تنظیمات هدر
            CsvStyleDto headerStyle = new CsvStyleDto();
            headerStyle.setAlignment(Element.ALIGN_CENTER);
            headerStyle.setBackgroundColor(BaseColor.BLUE);
            headerStyle.setFontColor(BaseColor.WHITE);
            headerStyle.setFontSize(20);
            headerStyle.setFontPath("static/pdf/font/arial/arial.ttf");

            // csv تنظیمات بادی
            CsvStyleDto bodyStyle = new CsvStyleDto();
            bodyStyle.setAlignment(Element.ALIGN_CENTER);
            bodyStyle.setBackgroundColor(BaseColor.WHITE);
            bodyStyle.setFontColor(BaseColor.BLACK);
            bodyStyle.setFontSize(15);
            bodyStyle.setBorderColor(BaseColor.BLACK);
            bodyStyle.setFontPath("static/pdf/font/arial/arial.ttf");

            //تنظمات سند
            CsvDto csvDto = new CsvDto();
            csvDto.setSeparator(";");
            csvDto.setHasHeader(true);
            csvDto.setDocumentRightToLeft(true);
            csvDto.setByteArrayInputStream(new ByteArrayInputStream(FileUtils.readFileToByteArray(file)));
            csvDto.setCsvHeaderStyleDto(headerStyle);
            csvDto.setCsvBodyStyleDto(bodyStyle);
            csvDto.setPageSize(PageSize.A4.rotate());

            try (OutputStream outputStream = new FileOutputStream("test_converted/PdfToolsUnitTest_generateFromCsvTest.pdf")) {
                PdfTools.generateFromCsv(csvDto).writeTo(outputStream);
            }

        } catch (Exception ex) {
            fail(ex.toString());
        }
    }

    @Order(3)
    @Test
    void generateFromExcelTest() throws Exception {
        ClassLoader classLoader = getClass().getClassLoader();
        File file = new File(classLoader.getResource("static/pdf/testExcel/test.xlsx").getPath());

        try (OutputStream outputStream = new FileOutputStream("test_converted/PdfToolsUnitTest_generateFromExcelTest.pdf")) {
            PdfTools.generateFromExcel(new ByteArrayInputStream(FileUtils.readFileToByteArray(file)), PdfPageSize.A3, PdfOrientation.PORTRAIT).writeTo(outputStream);
        }

    }
}

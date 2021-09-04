package com.motaharinia.msconverterutility.excel;

import com.motaharinia.msconverterutility.excel.dto.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;

import java.awt.*;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.fail;

/**
 * @author https://github.com/motaharinia<br>
 * کلاس تست ExcelTools
 */
@TestMethodOrder(MethodOrderer.OrderAnnotation.class)
class ExcelToolsUnitTest {


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
    void generateTest() {
        try {

            //تنظیمات لوکال که راست به چپ باشد یا خیر
            boolean sheetRightToLeft = true;

            //فرمت کننده برای فیلدهای boolean
            HashMap<Object, Object> formatterMap = new HashMap<>();
            formatterMap.put(true, "بلی");
            formatterMap.put(false, "خیر");

            ExcelStyleDto excelStyleDto;

            //عنوان سربرگ اکسل
            excelStyleDto = new ExcelStyleDto(HorizontalAlignment.CENTER, "Tahoma",true, Color.WHITE, new Color(198, 0, 199), BorderStyle.THIN, Color.BLACK,"General");
            ExcelCaptionDto captionDto = new ExcelCaptionDto("گزارش تیرماه اطلاعات کاربران", excelStyleDto);

            //عناوین ستونهای اکسل
            excelStyleDto = new ExcelStyleDto(HorizontalAlignment.CENTER, "Tahoma",true, Color.BLACK, new Color(49, 204, 206), BorderStyle.THIN, Color.BLACK,"General");
            List<ExcelColumnHeaderDto> columnHeaderList = new ArrayList<>();
            columnHeaderList.add(new ExcelColumnHeaderDto("نام", excelStyleDto));
            columnHeaderList.add(new ExcelColumnHeaderDto("نام خانوادگی", excelStyleDto));
            columnHeaderList.add(new ExcelColumnHeaderDto("دریافت خبرنامه؟", excelStyleDto));
            columnHeaderList.add(new ExcelColumnHeaderDto("امتیاز", excelStyleDto));
            columnHeaderList.add(new ExcelColumnHeaderDto("تعداد گردش", excelStyleDto));
            columnHeaderList.add(new ExcelColumnHeaderDto("معدل", excelStyleDto));
            columnHeaderList.add(new ExcelColumnHeaderDto("ضریب محاسبه", excelStyleDto));
            columnHeaderList.add(new ExcelColumnHeaderDto("تعداد مراجعه", excelStyleDto));
            columnHeaderList.add(new ExcelColumnHeaderDto("مبلغ موجودی", excelStyleDto));

            //تنظیمات ستونهای اکسل
            excelStyleDto = new ExcelStyleDto(HorizontalAlignment.CENTER, "Tahoma",false, Color.BLACK, Color.WHITE, BorderStyle.THIN, Color.BLACK,"General");
            ExcelStyleDto excelStyleDtoNumeric = new ExcelStyleDto(HorizontalAlignment.CENTER, "Tahoma",false, Color.BLUE, Color.WHITE, BorderStyle.THIN, Color.BLACK,"#,##0");
            ExcelStyleDto excelStyleDtoNumericFloat = new ExcelStyleDto(HorizontalAlignment.CENTER, "Tahoma",false, Color.BLUE, Color.WHITE, BorderStyle.THIN, Color.BLACK,"#,##0.00");
            List<ExcelColumnDto> columnList = new ArrayList<>();
            columnList.add(new ExcelColumnDto( null, excelStyleDto));
            columnList.add(new ExcelColumnDto( null, excelStyleDto));
            columnList.add(new ExcelColumnDto( formatterMap, excelStyleDto));
            columnList.add(new ExcelColumnDto( null, excelStyleDtoNumeric));
            columnList.add(new ExcelColumnDto( null, excelStyleDtoNumeric));
            columnList.add(new ExcelColumnDto( null, excelStyleDtoNumericFloat));
            columnList.add(new ExcelColumnDto( null, excelStyleDtoNumericFloat));
            columnList.add(new ExcelColumnDto( null, excelStyleDtoNumeric));
            columnList.add(new ExcelColumnDto( null, excelStyleDtoNumeric));

            //سطرهای اکسل
            List<Object[]> rowList = new ArrayList<>();
            rowList.add(new Object[]{"مصطفی", "مطهری نیا", true, 10000, 1000000L, 1.5f, 1.25d, BigInteger.valueOf(1000000000L), BigDecimal.valueOf(150000)});
            rowList.add(new Object[]{"احمد", "کریمی راد", false, 20000, 2000000L, 2.5f, 2.25d, BigInteger.valueOf(2000000000L), BigDecimal.valueOf(250000)});

            //تولید اکسل
            ExcelDto excelDto = new ExcelDto("اطلاعات کاربران سامانه", sheetRightToLeft, captionDto, columnHeaderList, columnList, rowList);
            XSSFWorkbook workbook = ExcelTools.generate(excelDto);
            FileOutputStream fileOutputStream = new FileOutputStream("test_converted/ExcelToolsUnitTest_generateTest.xlsx");
            workbook.write(fileOutputStream);
            assertThat(workbook).isNotNull();
        } catch (Exception ex) {
            fail(ex.toString());
        }
    }


}

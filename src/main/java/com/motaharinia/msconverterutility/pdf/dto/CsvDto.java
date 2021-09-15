package com.motaharinia.msconverterutility.pdf.dto;

import com.itextpdf.text.Rectangle;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.ByteArrayInputStream;

/**
 * @author https://github.com/motaharinia<br>
 * csv کلاس مدل تنظمیات
 */

@AllArgsConstructor
@NoArgsConstructor
@Data
public class CsvDto {

    /**
     * pdf جهت صفحه
     */
    Boolean documentRightToLeft = true;
    /**
     * آیا هدر دارد یا نه
     */
    Boolean hasHeader = true;
    /**
     * csv جدا کننده
     */
    String separator = ",";
    /**
     * csv بایت ورودی
     */
    ByteArrayInputStream byteArrayInputStream;
    /**
     * csv ظاهر هدر
     */
    CsvStyleDto csvHeaderStyleDto;
    /**
     * csv ظاهر بادی
     */
    CsvStyleDto csvBodyStyleDto;
    /**
     * اندازه صفحه
     */
    Rectangle pageSize;

}

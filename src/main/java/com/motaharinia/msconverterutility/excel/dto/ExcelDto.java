package com.motaharinia.msconverterutility.excel.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * @author https://github.com/motaharinia<br>
 * کلاس مدل تنظمیات اکسل
 */

@AllArgsConstructor
@NoArgsConstructor
@Data
public class ExcelDto {

    /**
     * عنوان صفحه اکسل
     */
    String sheetTitle = "sheet1";
    /**
     *جهت صفحه اکسل
     */
    Boolean sheetRightToLeft = true;
    /**
     * عنوان سربرگ اکسل
     */
    private ExcelCaptionDto captionDto;
    /**
     * لیستی از عناوین ستونهای اکسل را در خود دارد
     */
    private List<ExcelColumnHeaderDto> columnHeaderList;
    /**
     * لیستی از تنظمیات ستونهای اکسل را در خود دارد
     */
    private List<ExcelColumnDto> columnList;
    /**
     * داده های سطرها
     */
    private List<Object[]> rowList;

}

package com.motaharinia.msconverterutility.excel.dto;

import java.util.List;

/**
 * @author https://github.com/motaharinia<br>
 * کلاس اینترفیس مدل اطلاعات و تنظیمات تولید اکسل
 */
public interface CustomExcelDto {
    /**
     * @return خروجی: عنوان صفحه اکسل
     */
    String getSheetTitle();

    /**
     * @return خروجی: جهت راست به چپ صفحه اکسل
     */
    Boolean getSheetRightToLeft();

    /**
     * @return خروجی:  عنوان سربرگ اکسل
     */
    CustomExcelCaptionDto getCaptionDto();

    /**
     * @return خروجی:  لیستی از عناوین ستونهای اکسل را در خود دارد
     */
    List<CustomExcelColumnHeaderDto> getColumnHeaderList();

    /**
     * @return خروجی:  لیستی از تنظمیات ستونهای اکسل را در خود دارد
     */
    List<CustomExcelColumnDto> getColumnList();

    /**
     * @return خروجی: لیست سطرهای داده
     */
    List<Object[]> getRowList();
}

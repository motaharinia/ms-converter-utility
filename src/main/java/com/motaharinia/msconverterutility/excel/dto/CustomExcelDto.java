package com.motaharinia.msconverterutility.excel.dto;

import java.util.List;

public interface CustomExcelDto {
    /**
     * عنوان صفحه اکسل
     */
    String getSheetTitle();
    /**
     *جهت صفحه اکسل
     */
    Boolean getSheetRightToLeft();
    /**
     * عنوان سربرگ اکسل
     */
    CustomExcelCaptionDto getCaptionDto();
    /**
     * لیستی از عناوین ستونهای اکسل را در خود دارد
     */
    List<CustomExcelColumnHeaderDto> getColumnHeaderList();
    /**
     * لیستی از تنظمیات ستونهای اکسل را در خود دارد
     */
    List<CustomExcelColumnDto> getColumnList();
    /**
     * داده های سطرها
     */
    List<Object[]> getRowList();
}

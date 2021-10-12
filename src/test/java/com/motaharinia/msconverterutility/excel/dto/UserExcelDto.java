package com.motaharinia.msconverterutility.excel.dto;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.awt.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class UserExcelDto implements CustomExcelDto {

    /**
     * شیی تنظیمات ظاهری
     */
    private CustomExcelStyleDto customExcelStyleDto;

    /**
     *  داده های سطرها
     */
    private List<Object[]> rowList = new ArrayList<>();


    /**
     * متد سازنده
     * @param rowList داده های سطرها
     */
    public UserExcelDto(List<Object[]> rowList) {
        this.rowList = rowList;
    }

    /**
     * عنوان صفحه اکسل
     */
    @Override
    public String getSheetTitle() {
        return "اطلاعات کاربران سامانه";
    }

    /**
     * جهت صفحه اکسل
     */
    @Override
    public Boolean getSheetRightToLeft() {
        return true;
    }

    /**
     * عنوان سربرگ اکسل
     */
    @Override
    public CustomExcelCaptionDto getCaptionDto() {
        //عنوان سربرگ اکسل
        customExcelStyleDto = new CustomExcelStyleDto(HorizontalAlignment.CENTER, "Tahoma",true, Color.WHITE, new Color(198, 0, 199), BorderStyle.THIN, Color.BLACK,"General");
        return new CustomExcelCaptionDto("گزارش تیرماه اطلاعات کاربران", customExcelStyleDto);
    }

    /**
     * لیستی از عناوین ستونهای اکسل را در خود دارد
     */
    @Override
    public List<CustomExcelColumnHeaderDto> getColumnHeaderList() {
        //عناوین ستونهای اکسل
        customExcelStyleDto = new CustomExcelStyleDto(HorizontalAlignment.CENTER, "Tahoma",true, Color.BLACK, new Color(49, 204, 206), BorderStyle.THIN, Color.BLACK,"General");
        List<CustomExcelColumnHeaderDto> columnHeaderList = new ArrayList<>();
        columnHeaderList.add(new CustomExcelColumnHeaderDto("نام", customExcelStyleDto));
        columnHeaderList.add(new CustomExcelColumnHeaderDto("نام خانوادگی", customExcelStyleDto));
        columnHeaderList.add(new CustomExcelColumnHeaderDto("دریافت خبرنامه؟", customExcelStyleDto));
        columnHeaderList.add(new CustomExcelColumnHeaderDto("امتیاز", customExcelStyleDto));
        columnHeaderList.add(new CustomExcelColumnHeaderDto("تعداد گردش", customExcelStyleDto));
        columnHeaderList.add(new CustomExcelColumnHeaderDto("معدل", customExcelStyleDto));
        columnHeaderList.add(new CustomExcelColumnHeaderDto("ضریب محاسبه", customExcelStyleDto));
        columnHeaderList.add(new CustomExcelColumnHeaderDto("تعداد مراجعه", customExcelStyleDto));
        columnHeaderList.add(new CustomExcelColumnHeaderDto("مبلغ موجودی", customExcelStyleDto));
        return columnHeaderList;
    }

    /**
     * لیستی از تنظمیات ستونهای اکسل را در خود دارد
     */
    @Override
    public List<CustomExcelColumnDto> getColumnList() {
        //فرمت کننده برای فیلدهای boolean
        HashMap<Object, Object> formatterMap = new HashMap<>();
        formatterMap.put(true, "بلی");
        formatterMap.put(false, "خیر");

        //تنظیمات ستونهای اکسل
        customExcelStyleDto = new CustomExcelStyleDto(HorizontalAlignment.CENTER, "Tahoma",false, Color.BLACK, Color.WHITE, BorderStyle.THIN, Color.BLACK,"General");
        CustomExcelStyleDto customExcelStyleDtoNumeric = new CustomExcelStyleDto(HorizontalAlignment.CENTER, "Tahoma",false, Color.BLUE, Color.WHITE, BorderStyle.THIN, Color.BLACK,"#,##0");
        CustomExcelStyleDto customExcelStyleDtoNumericFloat = new CustomExcelStyleDto(HorizontalAlignment.CENTER, "Tahoma",false, Color.BLUE, Color.WHITE, BorderStyle.THIN, Color.BLACK,"#,##0.00");
        List<CustomExcelColumnDto> columnList = new ArrayList<>();
        columnList.add(new CustomExcelColumnDto( null, customExcelStyleDto));
        columnList.add(new CustomExcelColumnDto( null, customExcelStyleDto));
        columnList.add(new CustomExcelColumnDto( formatterMap, customExcelStyleDto));
        columnList.add(new CustomExcelColumnDto( null, customExcelStyleDtoNumeric));
        columnList.add(new CustomExcelColumnDto( null, customExcelStyleDtoNumeric));
        columnList.add(new CustomExcelColumnDto( null, customExcelStyleDtoNumericFloat));
        columnList.add(new CustomExcelColumnDto( null, customExcelStyleDtoNumericFloat));
        columnList.add(new CustomExcelColumnDto( null, customExcelStyleDtoNumeric));
        columnList.add(new CustomExcelColumnDto( null, customExcelStyleDtoNumeric));
        return columnList;
    }

    /**
     * داده های سطرها
     */
    @Override
    public List<Object[]> getRowList() {
        return rowList;
    }
}

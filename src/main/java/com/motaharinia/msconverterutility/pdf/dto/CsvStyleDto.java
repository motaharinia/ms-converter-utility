package com.motaharinia.msconverterutility.pdf.dto;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;


import java.io.Serializable;

/**
 * @author https://github.com/motaharinia<br>
 * csv کلاس تنظیمات ظاهری
 */

@AllArgsConstructor
@NoArgsConstructor
@Data
public class CsvStyleDto implements Serializable {
    /**
     * جهت
     */
    private int alignment = Element.ALIGN_CENTER;
    /**
     * مسیر فونت
     */
    private String fontPath = "";
    /**
     * برجسته بودن قلم
     */
    private int fontStyle = Font.NORMAL;
    /**
     * رنگ قلم
     */
    private BaseColor fontColor = BaseColor.BLACK;
    /**
     * رنگ قلم
     */
    private int fontSize = 20;
    /**
     * رنگ پس زمینه
     */
    private BaseColor backgroundColor = BaseColor.BLUE;
    /**
     * رنگ جدول
     */
    private BaseColor borderColor = BaseColor.WHITE;
}

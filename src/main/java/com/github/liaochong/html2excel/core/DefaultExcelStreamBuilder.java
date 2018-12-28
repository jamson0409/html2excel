package com.github.liaochong.html2excel.core;

import com.github.liaochong.html2excel.core.annotation.ExcelColumn;
import com.github.liaochong.html2excel.core.annotation.ExcelTable;
import com.github.liaochong.html2excel.core.annotation.ExcludeColumn;
import com.github.liaochong.html2excel.core.cache.Cache;
import com.github.liaochong.html2excel.core.cache.DefaultCache;
import com.github.liaochong.html2excel.core.parser.Td;
import com.github.liaochong.html2excel.core.parser.Tr;
import com.github.liaochong.html2excel.core.reflect.ClassFieldContainer;
import com.github.liaochong.html2excel.utils.ReflectUtil;
import com.github.liaochong.html2excel.utils.StringUtil;
import com.github.liaochong.html2excel.utils.TdUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CompletableFuture;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class DefaultExcelStreamBuilder {

    private static final Cache<String, DateTimeFormatter> DATETIME_FORMATTER_CONTAINER = new DefaultCache<>();

    /**
     * 标题
     */
    private List<String> titles;
    /**
     * sheetName
     */
    private String sheetName;
    /**
     * 字段展示顺序
     */
    private List<String> fieldDisplayOrder;

    private BlockingQueue<Tr> blockingQueue = new ArrayBlockingQueue<>(10000);

    private DefaultExcelStreamBuilder() {
    }

    public static DefaultExcelStreamBuilder getInstance() {
        return new DefaultExcelStreamBuilder();
    }

    public DefaultExcelStreamBuilder titles(List<String> titles) {
        this.titles = titles;
        return this;
    }

    public DefaultExcelStreamBuilder sheetName(String sheetName) {
        this.sheetName = Objects.isNull(sheetName) ? "sheet" : sheetName;
        return this;
    }

    public DefaultExcelStreamBuilder fieldDisplayOrder(List<String> fieldDisplayOrder) {
        this.fieldDisplayOrder = fieldDisplayOrder;
        return this;
    }

    public Workbook build(List<?> data) throws Exception {
        log.info("开始构建");
        long startTime = System.currentTimeMillis();
        if (Objects.isNull(data) || data.isEmpty()) {
            log.info("No valid data exists");
            return new HtmlToExcelFactory().build(Collections.emptyList());
        }
        Optional<?> findResult = data.stream().filter(Objects::nonNull).findFirst();
        if (!findResult.isPresent()) {
            log.info("No valid data exists");
            return new HtmlToExcelFactory().build(Collections.emptyList());
        }
        ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(findResult.get().getClass());
        List<Field> sortedFields = getSortedFieldsAndSetTitles(classFieldContainer);

        if (sortedFields.isEmpty()) {
            log.info("The specified field mapping does not exist");
            return new HtmlToExcelFactory().build(Collections.emptyList());
        }
        CompletableFuture<Workbook> workbookCompletableFuture = CompletableFuture.supplyAsync(() -> {
            try {
                return new HtmlToExcelStreamFactory(blockingQueue).build(sheetName);
            } catch (Exception e) {
                e.printStackTrace();
            }
            return null;
        });
        renderContent(data, sortedFields);
        Workbook workbook = workbookCompletableFuture.get();
        log.info("构建耗时：{}", System.currentTimeMillis() - startTime);
        return workbook;
    }

    /**
     * 获取排序后字段并设置标题
     *
     * @param classFieldContainer classFieldContainer
     * @return Field
     */
    private List<Field> getSortedFieldsAndSetTitles(ClassFieldContainer classFieldContainer) {
        ExcelTable excelTable = classFieldContainer.getClazz().getAnnotation(ExcelTable.class);
        List<String> titles = new ArrayList<>();
        List<Field> sortedFields;
        if (Objects.nonNull(excelTable)) {
            boolean excludeParent = excelTable.excludeParent();
            if (excludeParent) {
                sortedFields = classFieldContainer.getDeclaredFields();
            } else {
                sortedFields = classFieldContainer.getFields();
            }
            sortedFields = sortedFields.stream()
                    .filter(field -> !field.isAnnotationPresent(ExcludeColumn.class))
                    .sorted((field1, field2) -> {
                        ExcelColumn excelColumn1 = field1.getAnnotation(ExcelColumn.class);
                        ExcelColumn excelColumn2 = field2.getAnnotation(ExcelColumn.class);
                        if (Objects.isNull(excelColumn1) && Objects.isNull(excelColumn2)) {
                            return 0;
                        }
                        int defaultOrder = 0;
                        int order1 = defaultOrder;
                        if (Objects.nonNull(excelColumn1)) {
                            order1 = excelColumn1.order();
                        }
                        int order2 = defaultOrder;
                        if (Objects.nonNull(excelColumn2)) {
                            order2 = excelColumn2.order();
                        }
                        if (order1 == order2) {
                            return 0;
                        }
                        return order1 > order2 ? 1 : -1;
                    })
                    .peek(field -> {
                        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                        if (Objects.isNull(excelColumn)) {
                            titles.add(null);
                        } else {
                            titles.add(excelColumn.title());
                        }
                    })
                    .collect(Collectors.toList());
        } else {
            List<Field> excelColumnFields = classFieldContainer.getFieldsByAnnotation(ExcelColumn.class);
            if (excelColumnFields.isEmpty()) {
                if (Objects.isNull(fieldDisplayOrder) || fieldDisplayOrder.isEmpty()) {
                    throw new IllegalArgumentException("FieldDisplayOrder is necessary");
                }
                this.selfAdaption();
                return fieldDisplayOrder.stream()
                        .map(classFieldContainer::getFieldByName)
                        .collect(Collectors.toList());
            }
            sortedFields = excelColumnFields.stream()
                    .filter(field -> !field.isAnnotationPresent(ExcludeColumn.class))
                    .sorted((field1, field2) -> {
                        int order1 = field1.getAnnotation(ExcelColumn.class).order();
                        int order2 = field2.getAnnotation(ExcelColumn.class).order();
                        if (order1 == order2) {
                            return 0;
                        }
                        return order1 > order2 ? 1 : -1;
                    }).peek(field -> {
                        String title = field.getAnnotation(ExcelColumn.class).title();
                        titles.add(title);
                    })
                    .collect(Collectors.toList());
        }

        boolean hasTitle = titles.stream().anyMatch(StringUtil::isNotBlank);
        if (hasTitle) {
            this.titles = titles;
        }
        return sortedFields;
    }

    /**
     * 展示字段order与标题title长度一致性自适应
     */
    private void selfAdaption() {
        if (Objects.isNull(titles) || titles.isEmpty()) {
            return;
        }
        if (fieldDisplayOrder.size() < titles.size()) {
            for (int i = 0, size = titles.size() - fieldDisplayOrder.size(); i < size; i++) {
                fieldDisplayOrder.add(null);
            }
        } else {
            for (int i = 0, size = fieldDisplayOrder.size() - titles.size(); i < size; i++) {
                titles.add(null);
            }
        }
    }

    /**
     * 获取并且转换字段值
     *
     * @param data  数据
     * @param field 对应字段
     * @return 结果
     */
    private Object getAndConvertFieldValue(Object data, Field field) {
        Object result = ReflectUtil.getFieldValue(data, field);
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        if (Objects.isNull(excelColumn) || Objects.isNull(result)) {
            return result;
        }
        // 时间格式化
        String dateFormatPattern = excelColumn.dateFormatPattern();
        if (StringUtil.isNotBlank(dateFormatPattern)) {
            Class<?> fieldType = field.getType();
            if (fieldType == LocalDateTime.class) {
                LocalDateTime localDateTime = (LocalDateTime) result;
                DateTimeFormatter formatter = getDateTimeFormatter(dateFormatPattern);
                return formatter.format(localDateTime);
            } else if (fieldType == LocalDate.class) {
                LocalDate localDate = (LocalDate) result;
                DateTimeFormatter formatter = getDateTimeFormatter(dateFormatPattern);
                return formatter.format(localDate);
            } else if (fieldType == Date.class) {
                Date date = (Date) result;
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dateFormatPattern);
                return simpleDateFormat.format(date);
            }
        }
        return result;
    }

    /**
     * 获取时间格式化
     *
     * @param dateFormat 时间格式化
     * @return DateTimeFormatter
     */
    private DateTimeFormatter getDateTimeFormatter(String dateFormat) {
        DateTimeFormatter formatter = DATETIME_FORMATTER_CONTAINER.get(dateFormat);
        if (Objects.isNull(formatter)) {
            formatter = DateTimeFormatter.ofPattern(dateFormat);
            DATETIME_FORMATTER_CONTAINER.cache(dateFormat, formatter);
        }
        return formatter;
    }

    /**
     * 获取需要被渲染的内容
     *
     * @param data         数据集合
     * @param sortedFields 排序字段
     */
    private void renderContent(List<?> data, List<Field> sortedFields) throws Exception {
        Map<String, String> commonStyle = new HashMap<>();
        commonStyle.put("border-bottom-style", "thin");
        commonStyle.put("border-left-style", "thin");
        commonStyle.put("border-right-style", "thin");

        boolean hasTitles = Objects.nonNull(titles) && !titles.isEmpty();
        if (hasTitles) {
            Tr tr = getThead(commonStyle);
            blockingQueue.put(tr);
        }
        Map<String, String> oddTdStyle = new HashMap<>(commonStyle);
        oddTdStyle.put("background-color", "#f6f8fa");
        // 偏移量
        int shift = hasTitles ? 1 : 0;
        IntStream.range(0, data.size()).parallel().forEach(index -> {
            List<Object> dataList = sortedFields.stream()
                    .map(field -> this.getAndConvertFieldValue(data.get(index), field))
                    .collect(Collectors.toList());
            int trIndex = index + shift;
            Tr tr = new Tr(trIndex);
            Map<Integer, Integer> colMaxWidthMap = new HashMap<>(dataList.size());
            tr.setColWidthMap(colMaxWidthMap);
            Map<String, String> tdStyle = tr.getIndex() % 2 == 0 ? commonStyle : oddTdStyle;
            List<Td> tdList = IntStream.range(0, dataList.size()).mapToObj(i -> {
                Td td = new Td();
                td.setRow(trIndex);
                td.setRowBound(trIndex);
                td.setCol(i);
                td.setColBound(i);
                td.setContent(Objects.isNull(dataList.get(i)) ? null : String.valueOf(dataList.get(i)));
                td.setStyle(tdStyle);
                tr.getColWidthMap().put(i, TdUtil.getStringWidth(td.getContent()));
                return td;
            }).collect(Collectors.toList());
            tr.setTdList(tdList);
            try {
                blockingQueue.put(tr);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        });
    }

    /**
     * 获取thead
     *
     * @param commonStyle 公共style
     * @return tr
     */
    private Tr getThead(Map<String, String> commonStyle) {
        Tr tr = new Tr(0);
        Map<String, String> thStyle = new HashMap<>();
        thStyle.put("font-weight", "bold");
        thStyle.put("font-size", "14");
        thStyle.put("text-align", "center");
        thStyle.put("vertical-align", "center");
        thStyle.putAll(commonStyle);

        Map<Integer, Integer> colMaxWidthMap = new HashMap<>(titles.size());
        tr.setColWidthMap(colMaxWidthMap);

        List<Td> ths = IntStream.range(0, titles.size()).mapToObj(index -> {
            Td td = new Td();
            td.setTh(true);
            td.setRow(0);
            td.setRowBound(0);
            td.setCol(index);
            td.setColBound(index);
            td.setContent(titles.get(index));
            td.setStyle(thStyle);
            tr.getColWidthMap().put(index, TdUtil.getStringWidth(td.getContent()));
            return td;
        }).collect(Collectors.toList());
        tr.setTdList(ths);
        return tr;
    }
}

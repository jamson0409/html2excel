package com.github.liaochong.html2excel.core;

import com.github.liaochong.html2excel.core.parser.HtmlTableParser;
import com.github.liaochong.html2excel.core.parser.Table;
import com.github.liaochong.html2excel.core.parser.Td;
import com.github.liaochong.html2excel.core.parser.Tr;
import com.github.liaochong.html2excel.core.style.BackgroundStyle;
import com.github.liaochong.html2excel.core.style.BorderStyle;
import com.github.liaochong.html2excel.core.style.FontStyle;
import com.github.liaochong.html2excel.core.style.TdDefaultCellStyle;
import com.github.liaochong.html2excel.core.style.TextAlignStyle;
import com.github.liaochong.html2excel.core.style.ThDefaultCellStyle;
import com.github.liaochong.html2excel.exception.UnsupportedWorkbookTypeException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.EnumMap;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class HtmlToExcelStreamFactory {

    private static final long DEFAULT_WAIT_TIME = 200L;

    private long waitTime = DEFAULT_WAIT_TIME;
    /**
     * excel workbook
     */
    private Workbook workbook;

    private BlockingQueue<Tr> blockingQueue;
    /**
     * 样式容器
     */
    private Map<HtmlTableParser.TableTag, CellStyle> defaultCellStyleMap;
    /**
     * 单元格样式映射
     */
    private Map<Map<String, String>, CellStyle> cellStyleMap = new HashMap<>();
    /**
     * 每行的单元格最大高度map
     */
    private Map<Integer, Short> maxTdHeightMap;
    /**
     * 字体map
     */
    private Map<String, Font> fontMap = new HashMap<>();
    /**
     * 是否使用默认样式
     */
    private boolean useDefaultStyle;
    /**
     * 自定义颜色索引
     */
    private AtomicInteger colorIndex = new AtomicInteger(56);

    public HtmlToExcelStreamFactory(BlockingQueue<Tr> blockingQueue) {
        this.blockingQueue = blockingQueue;
    }

    /**
     * 设置workbook类型
     *
     * @param workbookType 工作簿类型
     * @return HtmlToExcelFactory
     */
    public HtmlToExcelStreamFactory workbookType(WorkbookType workbookType) {
        if (Objects.isNull(workbookType)) {
            throw new IllegalArgumentException("WorkbookType must be specified,or remove this method, use the default workbookType");
        }
        switch (workbookType) {
            case XLS:
                workbook = new HSSFWorkbook();
                break;
            case XLSX:
                workbook = new XSSFWorkbook();
                break;
            case SXLSX:
                throw new UnsupportedWorkbookTypeException("SXSSFWorkbook is not supported at this version");
            default:
        }
        return this;
    }

    public HtmlToExcelStreamFactory waitTime(long waitTime) {
        this.waitTime = waitTime;
        return this;
    }

    /**
     * 设置使用默认样式
     *
     * @return HtmlToExcelFactory
     */
    public HtmlToExcelStreamFactory useDefaultStyle() {
        this.useDefaultStyle = true;
        return this;
    }

    public Workbook build(String sheetName) throws Exception {
        if (Objects.isNull(workbook)) {
            workbook = new XSSFWorkbook();
        }

        if (useDefaultStyle) {
            defaultCellStyleMap = new EnumMap<>(HtmlTableParser.TableTag.class);
            defaultCellStyleMap.put(HtmlTableParser.TableTag.th, new ThDefaultCellStyle().supply(workbook));
            defaultCellStyleMap.put(HtmlTableParser.TableTag.td, new TdDefaultCellStyle().supply(workbook));
        }

        Sheet sheet = workbook.createSheet(Objects.isNull(sheetName) ? "sheet" : sheetName);
        Tr tr = blockingQueue.poll(waitTime, TimeUnit.MILLISECONDS);
        while (Objects.nonNull(tr)) {
            log.info("处理行：{}", tr.getIndex());
            // 设置单元格样式
            tr.getTdList().forEach(td -> this.setCell(td, sheet));
            // 设置行高
//            this.setRowHeight(sheet, );

//            if (Objects.nonNull(freezePanes) && freezePanes.length > i) {
//                FreezePane freezePane = freezePanes[i];
//                if (Objects.isNull(freezePane)) {
//                    throw new IllegalStateException("FreezePane is null");
//                }
//                sheet.createFreezePane(freezePane.getColSplit(), freezePane.getRowSplit());
//            }
            tr = blockingQueue.poll(waitTime, TimeUnit.MILLISECONDS);
        }
        return workbook;
    }


    /**
     * 设置所有单元格，自适应列宽，单元格最大支持字符长度255
     */
    private void setTd(Table table, Sheet sheet) {
        for (int i = 0; i < table.getTrList().size(); i++) {
            Tr tr = table.getTrList().get(i);
            tr.getTdList().forEach(td -> this.setCell(td, sheet));
            table.getTrList().set(i, null);
        }

        table.getColMaxWidthMap().forEach((key, value) -> {
            int contentLength = value << 1;
            if (contentLength > 255) {
                contentLength = 255;
            }
            sheet.setColumnWidth(key, contentLength << 8);
        });
    }

    /**
     * 设置行高，最小12
     */
    private void setRowHeight(Sheet sheet, int rowSize) {
        for (int j = 0; j < rowSize; j++) {
            Row row = sheet.getRow(j);
            if (Objects.isNull(maxTdHeightMap.get(row.getRowNum()))) {
                row.setHeightInPoints(row.getHeightInPoints() + 5);
            } else {
                row.setHeightInPoints((short) (maxTdHeightMap.get(row.getRowNum()) + 5));
            }
        }
    }

    /**
     * 设置单元格
     *
     * @param td    单元格
     * @param sheet 单元格所在的sheet
     */
    private void setCell(Td td, Sheet sheet) {
        Row currentRow = sheet.getRow(td.getRow());
        if (Objects.isNull(currentRow)) {
            currentRow = sheet.createRow(td.getRow());
        }

        Cell cell = currentRow.getCell(td.getCol());
        if (Objects.isNull(cell)) {
            cell = currentRow.createCell(td.getCol());
        }
        cell.setCellValue(td.getContent());


        // 设置单元格样式
        for (int i = td.getRow(), rowBound = td.getRowBound(); i <= rowBound; i++) {
            Row row = sheet.getRow(i);
            if (Objects.isNull(row)) {
                row = sheet.createRow(i);
            }
            for (int j = td.getCol(), colBound = td.getColBound(); j <= colBound; j++) {
                cell = row.getCell(j);
                if (Objects.isNull(cell)) {
                    cell = row.createCell(j);
                }
                this.setCellStyle(row, cell, td);
            }
        }
        if (td.getColSpan() > 0 || td.getRowSpan() > 0) {
            sheet.addMergedRegion(new CellRangeAddress(td.getRow(), td.getRowBound(), td.getCol(), td.getColBound()));
        }
    }

    /**
     * 设置单元格样式
     *
     * @param cell 单元格
     * @param td   td单元格
     */
    private void setCellStyle(Row row, Cell cell, Td td) {
        if (useDefaultStyle) {
            if (td.isTh()) {
                cell.setCellStyle(defaultCellStyleMap.get(HtmlTableParser.TableTag.th));
            } else {
                cell.setCellStyle(defaultCellStyleMap.get(HtmlTableParser.TableTag.td));
            }
        } else {
            if (cellStyleMap.containsKey(td.getStyle())) {
                cell.setCellStyle(cellStyleMap.get(td.getStyle()));
                return;
            }
            CellStyle cellStyle = workbook.createCellStyle();
            // background-color
            BackgroundStyle.setBackgroundColor(workbook, cellStyle, td.getStyle(), colorIndex);
            // text-align
            TextAlignStyle.setTextAlign(cellStyle, td.getStyle());
            // border
            BorderStyle.setBorder(cellStyle, td.getStyle());
            // font
            FontStyle.setFont(workbook, row, cellStyle, td.getStyle(), fontMap, maxTdHeightMap);
            cell.setCellStyle(cellStyle);
            cellStyleMap.put(td.getStyle(), cellStyle);
        }
    }
}

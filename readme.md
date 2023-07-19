# 背景
丰景台支持图表的数据取数，其中透视表类型的图表日均取数量在 次/日。将数据库查询出的数据下载为透视表展现样式的excel。
#POI数据格式
POI支持的数据格式包括HSSFWorkbook，XSSFWorkbook和SXSSFWorkbook。
- HSSFWorkbook对应excel2003，基本弃用。

- XSSFWorkbook对应excel2007，最多可以导出104万行，不过存在一个问题—OOM内存溢出，原因是创建的book sheet row cell等是存在内存的并没有持久化，大数据量的读写会出现OOM问题。

- SXSSFWorkbook是streaming版本的XSSFWorkbook,它只会保存最新的excel rows在内存里供查看，在此之前的excel rows都会被写入到硬盘里（Windows电脑的话，是写入到C盘根目录下的temp文件夹）。被写入到硬盘里的rows是不可见的/不可访问的。只有还保存在内存里的才可以被访问到。
#透视表下载方案
## 1. 使用excel的透视表格式
仅XSSFWorkbook支持将数据转换为excel的透视表格式（猜测：由于透视表需要对全量数据进行统计、汇总、计算等操作，所以只能读取内存中数据的SXSSFWorkbook格式不能支持透视表的创建，类似不支持公式计算。）其生成透视表的逻辑是将明细数据写入一个sheet页，基于该sheet页的数据创建透视表。DEMO步骤如下：
1. 创建sheet1，写入明细数据；
2. 以sheet1为数据源区域，即CellReference；
3. 创建透视表sheet，并在该sheet上创建povitTable；
4. 添加行标签、列标签、汇总计算值，设置格式等；
5. 导出文件。


```java
    public static void xssfWorkbooktest(){
        try {

            long startTime = System.currentTimeMillis();
            Object[][] data = new Object[][] { { "2021-08","23","01412262","丰景台demo数据集",1,4}, { "2021-10","23","01412262","丰景台demo数据集",2,5},
                    { "2021-08","23","01412262","ES数据",2,3}, { "2021-08","23","01412262","ES数据",2,4},
                    { "2021-08","23","01412262","yellowbrickdemo",23,2}};
            XSSFWorkbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet();
            Row row = sheet.createRow(0);
            String[] colNames = new String[]{"更新时间(年月)","自定义","创建者的Id","数据集名称","计算1","计算2"};
			//todo 写入数据
            for (int i = 0; i < 6; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(colNames[i]);
            }

            int num = 100000;

            // 数据透视表生产的起点单元格位置
            CellReference ptStartCell = new CellReference("A1");
            //根据你自己的要的表格的列的数量决定，
            AreaReference areaR=new AreaReference("A1:E"+num,org.apache.poi.ss.SpreadsheetVersion.EXCEL2007);
            XSSFSheet pivotSheet = workbook.createSheet("透视表");
            //从sheet1的选定数据范围内数据生成数据透视表
            XSSFPivotTable pivotTable = pivotSheet.createPivotTable(areaR, ptStartCell, sheet);

            CTPivotTableDefinition ctPivotTableDefinition=pivotTable.getCTPivotTableDefinition();
            //添加行标签
            pivotTable.addRowLabel(2);
            pivotTable.addRowLabel(3);

            //取消汇总
            ctPivotTableDefinition.setColGrandTotals(false);
            //非压缩，非大纲模式
            ctPivotTableDefinition.setCompact(false);
            ctPivotTableDefinition.setCompactData(false);
            ctPivotTableDefinition.setOutline(false);
            ctPivotTableDefinition.setOutlineData(false);

            //添加列标签
            pivotTable.addColLabel(1);
            pivotTable.addColLabel(0);

            //添加计算值
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 4,"计算1");
            pivotTable.addDataColumn(4,true);

            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 4,"计算2");
            pivotTable.addDataColumn(4,true);


            //取消行列总计
            ctPivotTableDefinition.setRowGrandTotals(false);
            ctPivotTableDefinition.setColGrandTotals(false);

            ctPivotTableDefinition.setEditData(true);

            for(CTDataField dataField : ctPivotTableDefinition.getDataFields().getDataFieldList()){

                //dataField.setShowDataAs(STShowDataAs.PERCENT_OF_COL);//不起作用问题
                //dataField.setNumFmtId(10);
            }

            XSSFPivotCacheDefinition cacheDefinition= pivotTable.getPivotCacheDefinition();
            CTPivotCacheDefinition cache = cacheDefinition.getCTPivotCacheDefinition();

            for (CTPivotField ctPivotField :  ctPivotTableDefinition.getPivotFields().getPivotFieldList()) {
                //取消分类汇总
                ctPivotField.setSubtotalTop(false);
                ctPivotField.setDefaultSubtotal(false);
                //表格模式
                ctPivotField.setCompact(false);
                ctPivotField.setOutline(false);

            }
            //合并单元格
            OutputStream os =new FileOutputStream("d:\\user\\01412262\\desktop\\pivotTable.xlsx");
            workbook.write(os);
            workbook.close();
            os.flush();
            os.close();
            long endTime2 = System.currentTimeMillis();
            System.out.println("写入耗时"+(endTime2-startTime)/1000 +"s");

        } catch (Exception e) {
            e.printStackTrace();
        }catch (Throwable e){
            System.out.println("f");
            e.printStackTrace();
        }
    }
```

- 存在的问题：大数据量创建透视表出现OOM
## 可支持百万行数据、行列数据排序、可自定义样式的透视表生成方案
为避免大数据量导出的OOM问题，必须使用SXSSFWorkbook格式完成透视表的创建。
SXSSFWorkbook是用来生成海量excel数据文件，主要原理是借助临时存储空间生成excel，SXSSFWorkbook专门处理大数据，对于大型excel的创建且不会内存溢出的，就只有SXSSFWorkbook了。它的原理很简单，用硬盘空间换内存（就像hashmap用空间换时间一样）。 SXSSFWorkbook是streaming版本的XSSFWorkbook,它只会保存最新的**randomAccessWindowSize** 行数据在内存里供查看，在此之前的excel rows都会被写入到硬盘里，保存为xml格式的文件（Windows电脑的话，是写入到C盘根目录下的temp文件夹）。被写入到硬盘里的rows是不可见的/不可访问的。只有还保存在内存里的才可以被访问到，源码如下。
```java
 public SXSSFRow createRow(int rownum) {
// ...
                if (this._randomAccessWindowSize >= 0 && this._rows.size() > this._randomAccessWindowSize) {
                    try {
                        this.flushRows(this._randomAccessWindowSize);
                    } catch (IOException var5) {
                        throw new RuntimeException(var5);
                    }
                }
				//...
    }

    public void flushRows(int remaining) throws IOException { //写入缓存
        while(this._rows.size() > remaining) {
            this.flushOneRow();
        }

        if (remaining == 0) {
            this.allFlushed = true;
        }

    }
```
### 目标
1、支持百万行数据导出生成透视表；
2、支持自定义行列数据排序；
3、支持自定义表格样式;
### 1. 数据查询
透视表数据查询sql如下：
基于数据查询引擎实现透视表数据的聚合计算和预排序。
```sql
select row1，row2，col1，col2，value1，value2 from tableA group by row1，row2，col1，col2 order by  row1，row2，col1，col2
```
数据查询结果List<List<String>> data;
![透视表源数据](http://osfp.sf-express.com/Public/Uploads/2023-07-10/64abb968ea224.png "透视表源数据")
### 2. 透视表数据建模
1. **透视表样例：**
![透视表样例](http://osfp.sf-express.com/Public/Uploads/2023-07-10/64abb9a1d3024.png "透视表样例")

1.  **透视表结构分析：**
![透视表结构](http://osfp.sf-express.com/Public/Uploads/2023-07-10/64abb9ddecdc4.jpg "透视表结构")

1. **透视表数据结构**
数据节点类
```java
   class DataItem {
        String name;//名称
        Integer index;// 位置
        Integer order;//排序规则
        String dataType;//数据类型
        Object cellType; //单元格数据类型
        XSSFCellStyle cellStyle; //单元格样式
   }
```
透视表类
```java
public class PivotTableV2 {

    private HeadNode rootColNode;

    private HeadNode rootRowNode;

    private List<DataItem> rows;//行

    private List<DataItem> cols;//列

    private List<DataItem> dataColumns;//数据

    private List<LinkedHashSet<String>> rowValues;//行值记录并去重

    private List<List<String>> rowOrderValues;//行值排序

    private Map<String, Integer> colIndexMap;//列值坐标

    private Integer colIndex;//顺序写起点

    private Integer rowIndex;

    private Integer accessDataWindowSize;//excel 可见行数

    private Integer rowNumBeWrite;// 不可见行数

    private Integer colNum;

    private Integer rowNum;

    private Integer dataNum;

    private Integer rowOffset;//行起始坐标偏移

    private SXSSFSheet sheet;

    private TableStyleConfig tableStyleConfig;//表样式配置

    private CellStyle dimNameCellStyle;//维度名样式

    private CellStyle indexNameCellStyle;//指标名样式

    private Integer mergeCell;//是否合并行头单元格

    private Integer mergeColumnCell;//是否合并列头单元格

    private Integer isFixedBody;//是否冻结行头

    private Integer isFixedHeader;//是否冻结列头
}
```

表头节点类
```java
    class HeadNode {
        private String name;
        private List<List<String>> values;//行带有列值和汇总值
        private Integer level;//层级
        private Integer range;
        private LinkedHashMap<String, HeadNode> childNodes;//下一层表头节点

        public HeadNode() {
            this.values = new ArrayList<>();
            this.level = -1;
            this.range = 1;
            this.name = "";
            this.childNodes = new LinkedHashMap<>();
		}
	}
```

 ### 3. 数据处理
1. 初始化数据节点
```java
    /**
     * 添加行
     *
     * @param name
     * @param index
     * @param dataType
     * @param orderType
     */
    public void addRow(String name, Integer index, String dataType, Integer orderType, Object cellType) {
        this.rows.add(new DataItem(name, index, dataType, orderType, cellType));
    }

    /**
     * 添加列
     *
     * @param name
     * @param index
     * @param dataType
     * @param orderType
     */
    public void addCol(String name, Integer index, String dataType, Integer orderType, Object cellType) {
        this.cols.add(new DataItem(name, index, dataType, orderType, cellType));
    }

    /**
     * 添加数据总计
     *
     * @param name
     * @param index
     * @param dataType
     * @param orderType
     */
    public void addDataColumn(String name, Integer index, String dataType, Integer orderType, Object cellType) {
        this.dataColumns.add(new DataItem(name, index, dataType, orderType, cellType));
    }
```
1. 读入数据并初始化表头节点
```java
    public void readRowData(List<String> data) {
        List<String> colValueData = new ArrayList<>();
        HeadNode tmp = rootColNode;//列值有序
        for (int i = 0; i < colNum; i++) {
            tmp = insertHeadNode(data.get(cols.get(i).index), tmp, i);
            colValueData.add(data.get(cols.get(i).index));
        }

        tmp = rootRowNode;//行值无序
        for (int i = 0; i < rowNum; i++) {
            if (rows.get(i).getOrder().equals(2) || rows.get(i).getOrder().equals(3)) {//1默认，2升序 3降序 4自定义
                rowValues.get(i).add(data.get(rows.get(i).index));
            }//排序
            tmp = insertHeadNode(data.get(rows.get(i).index), tmp, i);
        }
        for (int i = 0; i < dataNum; i++) {
            colValueData.add(data.get(dataColumns.get(i).index));
        }
        tmp.getValues().add(colValueData);

    }
	
    public HeadNode insertHeadNode(String name, HeadNode headNode, Integer level) {
        LinkedHashMap<String, HeadNode> childNodes = headNode.getChildNodes();
        if (childNodes != null && childNodes.containsKey(name)) {
            return childNodes.get(name);
        } else {
            HeadNode childNode = new HeadNode();
            childNode.setName(name);
            childNode.setLevel(level);
            headNode.getChildNodes().put(name, childNode);
            return childNode;
        }
    }
```
### 生成透视表

#### 生成透视表主流程
```java
    /**
     * 写数据
     *
     * @param sheet
     */
    public void writeTable(SXSSFSheet sheet) {
        sheet.setRandomAccessWindowSize(accessDataWindowSize);//可修改rowaccesswindowsize
        this.sheet = sheet;
        initTableStyle();//初始化样式
        writeHeadName(); //写角头单元格
        initAndWriteColHead(rootColNode, "");//写列头单元格，并记录排序
        orderRowValues();//行值排序
        writeRowAndValue(rootRowNode);//按顺序写行头和数据
        Integer fixRow = colNum + (dataNum == 0 ? 0 : 1);
        if(isFixedBody.equals(1) && isFixedHeader.equals(1)) {//冻结列头和行头
            sheet.createFreezePane(rowNum, fixRow);
        }else if(isFixedHeader.equals(1)){
            sheet.createFreezePane(0, fixRow);
        }else if(isFixedBody.equals(1)){
            sheet.createFreezePane(rowNum, 0);
        }
    }
```
#### 初始化表格样式
```java
    /**
     * 初始化表格样式
     */
    public void initTableStyle() {
        mergeCell = tableStyleConfig.getMergeCell();
        mergeColumnCell = tableStyleConfig.getMergeColumnCell();
        isFixedBody = tableStyleConfig.getIsFixedBody();
        isFixedHeader = tableStyleConfig.getIsFixedHeader();

        JSONObject dimHeadConfig = tableStyleConfig.getTableHeadStyle();
        JSONObject indexHeadConfig = tableStyleConfig.getTableHeadIndexStyle();
        JSONObject fillColorConfig = tableStyleConfig.getTableColor();
		...
	}
```

#### 深度遍历，写列头单元格，并记录坐标 colIndexMap
```java
    public Integer initAndWriteColHead(HeadNode headNode, String parentName) {
        LinkedHashMap<String, HeadNode> childNodes = headNode.getChildNodes();
        Integer dataNumFix = dataNum == 0 ? 1 : dataNum; //无指标值修正
        if ((childNodes == null || childNodes.size() == 0)) {
            Integer level = headNode.getLevel();
            if (level != -1) {
                //写最下层列头单元格
                for (int i = 0; i < dataNumFix; i++) {
                    writeCell(sheet, level, rowNum + colIndex * dataNumFix + i, headNode.getName(), cols.get(level).getCellStyle(), cols.get(level).getCellType());
                }
                //writeCell(sheet, level, rowNum + colIndex * dataNum, headNode.getName(), cols.get(level).getCellStyle(), cols.get(level).getCellType());
                if (dataNumFix > 1 && mergeCell.equals(1)) {//合并单元格
                    CellRangeAddress cellAddresses = new CellRangeAddress(headNode.getLevel(), headNode.getLevel(), rowNum + colIndex * dataNumFix, rowNum + (colIndex + 1) * dataNumFix - 1);
                    sheet.addMergedRegionUnsafe(cellAddresses);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, cellAddresses, sheet);
                    RegionUtil.setBorderTop(BorderStyle.THIN, cellAddresses, sheet);
                    RegionUtil.setBorderRight(BorderStyle.THIN, cellAddresses, sheet);
                    RegionUtil.setBorderBottom(BorderStyle.THIN, cellAddresses, sheet);
                }
            }//指标名单元格
            for (int i = 0; i < dataNum; i++) {
                writeCell(sheet, level + 1, rowNum + colIndex * dataNum + i, dataColumns.get(i).getName(), indexNameCellStyle, null);
            }
            colIndexMap.put(parentName + headNode.getName(), colIndex++);
            return headNode.getRange();
        }
        Integer range = 0;//单元格坐标范围
        for (HeadNode childNode : childNodes.values()) {
            range += initAndWriteColHead(childNode, parentName + headNode.getName());
        }
        headNode.setRange(range);
        //写单元格
        Integer level = headNode.getLevel();
        if (level != -1) {
            for (int i = 0; i < range; i++) {
                writeCell(sheet, level, rowNum + (colIndex - range) * dataNumFix + i, headNode.getName(), cols.get(level).getCellStyle(), cols.get(level).getCellType());
            }
            //writeCell(sheet, level, rowNum + (colIndex - range) * dataNum, headNode.getName(), cols.get(level).getCellStyle(), cols.get(level).getCellType());
            if(range > 1 && mergeCell.equals(1)) {
                CellRangeAddress cellAddresses = new CellRangeAddress(headNode.getLevel(), headNode.getLevel(), rowNum + (colIndex - range) * dataNumFix, rowNum + colIndex * dataNumFix - 1);
                sheet.addMergedRegionUnsafe(cellAddresses);
                RegionUtil.setBorderLeft(BorderStyle.THIN, cellAddresses, sheet);
                RegionUtil.setBorderTop(BorderStyle.THIN, cellAddresses, sheet);
                RegionUtil.setBorderRight(BorderStyle.THIN, cellAddresses, sheet);
                RegionUtil.setBorderBottom(BorderStyle.THIN, cellAddresses, sheet);
            }
        }
        return range;
    }
```
#### 深度遍历，按行值的排序，顺序写行头和数据
```java
    /**
     * 写行头和数据
     *
     * @param headNode
     * @return
     */
    public Integer writeRowAndValue(HeadNode headNode) {
        LinkedHashMap<String, HeadNode> childNodes = headNode.getChildNodes();
        if (childNodes == null || childNodes.size() == 0) {
            if (rowIndex >= accessDataWindowSize + rowNumBeWrite) {
                accessDataWindowSize += 10;
                sheet.setRandomAccessWindowSize(accessDataWindowSize);//动态调整
            }
            Integer level = headNode.getLevel();
            if (level != -1) {//没有行
                writeCell(sheet, rowIndex + rowOffset, level, headNode.getName(), rows.get(level).getCellStyle(), rows.get(level).getCellType());
            }
            List<List<String>> dataValues = headNode.getValues();
            Integer dataValuesNum = dataValues.size();
            for (int i = 0; i < dataValuesNum; i++) {
                List<String> rowData = dataValues.get(i);
                String key = "";
                for (int j = 0; j < colNum; j++) {
                    key += rowData.get(j);
                }
                Integer colIndex = colIndexMap.get(key);
                for (int j = 0; j < dataNum; j++) {
                    writeCell(sheet, rowIndex + rowOffset, rowNum + colIndex * dataNum + j, rowData.get(colNum + j), dataColumns.get(j).getCellStyle(), dataColumns.get(j).getCellType());
                }
            }
            rowIndex++;
            return headNode.getRange();
        }
        Integer range = 0;//单元格坐标范围
        Integer level = headNode.getLevel();
        List<String> rowValues = rowOrderValues.get(level + 1);
        Integer rowValueNum = rowValues.size();
        if (rowValueNum != 0) {//按行值排序顺序
            for (int i = 0; i < rowValueNum; i++) {
                HeadNode childNode = childNodes.get(rowValues.get(i));
                if (childNode != null) {
                    range += writeRowAndValue(childNode);//按顺序读取
                }
            }
        } else {//按读取数据顺序
            for (HeadNode childNode : childNodes.values()) {
                range += writeRowAndValue(childNode);
            }
        }
        headNode.setRange(range);
        //写单元格
        if (level != -1) {
            if(level == 0) {
                rowNumBeWrite = rowIndex;
            }
            for (int i = 0; i < range; i++) {
                writeCell(sheet, rowIndex + rowOffset - range + i, level, headNode.getName(), rows.get(level).getCellStyle(), rows.get(level).getCellType());
            }
            //writeCell(sheet, rowIndex + rowOffset - range, level, headNode.getName(), rows.get(level).getCellStyle(), rows.get(level).getCellType());
            if (range > 1 && mergeColumnCell.equals(1)) {
                CellRangeAddress cellAddresses = new CellRangeAddress(rowOffset + rowIndex - range, rowOffset + rowIndex - 1, headNode.getLevel(), headNode.getLevel());
                sheet.addMergedRegionUnsafe(cellAddresses);
            }
        }
        return range;
    }

```


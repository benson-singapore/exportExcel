# exportExcel
Excel tools，对excel封装让excel导出更简单！

 1、 新建excel导出。

``` java
Excel excel = new Excel(); //新建excel
ExcelSheet sheet = excel.createSheet(); //新建sheet
sheet.row(0).cell(2).cellValue("1");    //调用cellValue()，设置excel样式
sheet.row(1).cell(2).cellValue("2");
excel.saveExcel("c://test1.xlsx");      //存储excel
```

 2、 调用模板导出。

``` java
Excel excel = new Excel("c://test1.xlsx");
ExcelSheet sheet = excel.getSheet();         //默认获取第一个工作簿
sheet.row(0).cell(2).cellValue("111111111"); //设置excel value值
excel.saveExcel("c://test2.xlsx");
```
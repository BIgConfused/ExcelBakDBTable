# ExcelBakDBTable
从数据库导出指定表结构和数据到Excel中，表名为Excel名，字段为标题行，数据为对应字段的列
#
  修改jdbc.properties中的数据库连接信息，然后可修改POIOutputExcel中的runTest()方法，或者删掉重新自己调用OutputExcel执行
  传递的参数OutputExcel(String 表名,String Excel版本(xls[2003],xlsx[2007]),String 文件输出路径(如D:或者D:/都可以))

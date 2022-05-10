# operater-demanded
车间人员需求VBA自动生成

模块Semp函数用法

integer = Semp(range)

1. 函数根据所选择的单元格颜色及内容在人员数据库中查询所需人数。
2. 函数所使用的颜色信息应存储于Semp函数所在表（sheet）的第二行。
3. 函数所使用的人员数据信息应存储于相应名称的数据表（sheet）中，名称命名规则为：

	若：Semp函数所在表名称为		XXX
	则：数据库表名称应为				人员数据库（XXX）
	注：括号"（）"应为中文括号
	
4. 函数可以识别中文特殊工序（如：清场），但工序名必须存在于数据库表中，否则函数将返回0 。

%%% 版本更新记录 %%%
1.1.0 修复超长列字母溢出问题
1.1.2 修复小数人数异常取整问题

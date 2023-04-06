# 使用手册

有问题请在issue里提问

> 请注意, 开发者并没有义务回复您的问题. 您应该具备基本的提问技巧。
>
> 有关如何提问，请阅读[《提问的智慧》](https://github.com/ryanhanwu/How-To-Ask-Questions-The-Smart-Way/blob/main/README-zh_CN.md)
>

---

本项目遵循GPL 3.0开源协议。

#### 功能

通过从青年大学习后台中导出的已学习名单，对学生名单中的人进行筛选，最后生成一个*.xlsx文件。文件名为`f"[青年大学习学习记录]{out_date}"`（out_date为从青年大学习后台中导出文件的时间），其中有两个Sheet，第一个Sheet是未完成名单，第二个是各个支部的完成率。

#### 使用环境

- Python版本：Python 3.9.*

- 操作系统：Windows10

- Python第三方模块包：

  ```
  openpyxl 3.0.7
  ```

  

#### 使用方法

1. 在`student.csv`中提前设置好学生名单（建议为每个团支部赋予一个独一无二的编号或称呼）。
2. 将名单从后台中导出，另存为到项目根目录下`./Original Study Records`中，**保存成.csv格式文件（逗号分隔符），重命名为table_1.csv**。
3. 双击根目录下`launch.bat`，最新的学习记录保存在`./Study Records`文件夹中


#### FAQ
**Q: table_1.csv文件内容应该是怎样的格式？**

（感谢@Azurlane-a 提出的问题）

A: table_1.csv的内容格式说明如下

table_1.csv中的每一条信息，只需要能够包含个人的正确信息即可，必须包含的内容如下：

XXX支部，XXX（某人的姓名/工号/学号等）

注意，支部称呼最好唯一，**不同支部使用不同的称呼**。~~为了避免重名等因素导致的问题，**建议使用工号/学号等代替姓名**。~~ 

> 江西省青年大学习新增了 “学号/手机号” 一栏，现在无需使用 “工号/学号” 代替姓名。 “学号/手机号” 会出现在table_1.csv的 "备注" 一栏中。

一个合法的table_1.csv示例如下：

（其中 “地区四” 对应 “姓名” ，“备注” 对应 “学号/手机号” ）

**当你用excel打开时可能是这个样子**：

| 地区一 | 地区二 | 地区三 | 地区四 | 姓名 | 记录时间 | 备注         |
| ------ | ------ | ------ | ------ | ---- | -------- |------------|
| 省属本科院校团委 | XXXX大学团委 | XXX学院团委 | **22级法外狂徒专业** | **张三** | 2022-04-11 20:39:07.0 | 2022212000 |
| 省属本科院校团委 | XXXX大学团委 | XXX学院团委 | **22级混吃等死专业** | **李四** | 2022-04-11 20:43:07.0 | 2022212001 |

**当你用记事本打开时可能是这个样子**：

~~~
地区一, 地区二,地区三,地区四,姓名,记录时间,备注
省属本科院校团委,XXXX大学团委,XXX学院团委,22级法外狂徒专业,张三,2022-04-11 20:39:07.0,2022212000
省属本科院校团委,XXXX大学团委,XXX学院团委,22级混吃等死专业,李四,2022-04-11 20:43:07.0,2022212001
~~~


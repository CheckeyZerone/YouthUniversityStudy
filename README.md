# 使用手册

当前版本 `version 0.1.0`

有问题请在issue里提问

> 请注意, 开发者并没有义务回复您的问题. 您应该具备基本的提问技巧。
>
> 有关如何提问，请阅读[《提问的智慧》](https://github.com/ryanhanwu/How-To-Ask-Questions-The-Smart-Way/blob/main/README-zh_CN.md)
>

---

本项目遵循APACHE 2.0开源协议。

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



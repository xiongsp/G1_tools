# G1_tools

简单的为综测小组计算G1的工具

## 使用方法

1. 在Line3处指定成绩单excel文件路径
2. 修改target变量为统计G1的目标学号

```bash
# only tested in python3.10
pip install -r requirements.txt
python do.py
```
## Excel文件处理

Excel原始文件含有大量无用数据，应当处理为以下格式：

| 学号* | 姓名 | 课程名（课程代码）# | 学分* | 总评成绩* | 补考成绩* | 缓考成绩* |
| ----- | ---- | ------------------- | ----- | -------- | -------- | -------- |
|       |      |                     |       |          |          |          |

- *为必填项
- #为二选一项
- 姓名列可选，但是必须在第二列

## 计算规则

1. 缺考计为0分
2. 补考缺考按期末成绩计算
3. 总评成绩非数字者不计入G1
4. 补考成绩合格者按60计
5. 补考成绩仍不合格按照与期末成绩较大者计
6. 缓考成绩未出者不计入G1
7. 缓考成绩已出者按照缓考成绩计算（可能需要手动核实因公或因私缓考）

## 输出

输出信息包含可审计日志和G1计算结果，结果示例如下：

```
20211 71.94
20212 80.55555555555555555555555556
20224 80.98947368421052631578947368
20236 74.29885057471264367816091954
```
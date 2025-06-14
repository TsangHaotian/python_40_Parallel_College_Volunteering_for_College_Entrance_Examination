# python_40_Parallel_College_Volunteering_for_College_Entrance_Examination
3+2+1地区40个平行志愿高考报考模拟系统(基于2023年福建高考)

## 声明
### 根据政策最新可支持到2025年福建高考（如有报考政策变动，请以教育考试院公布的相关政策为准）
### 如有相同报考原则地区，该软件同样适用（本系统主要针对于滑档退档的规则解读）
#### 更新日期：2025年6月11日

## 系统简介
本系统是为3+2+1新高考地区(特别是福建省)设计的40个平行志愿填报模拟系统，包含三个主要模块：
1. 大学录取信息管理模块
2. 志愿填报模块
3. 志愿模拟录取模块

## 系统功能
- 大学专业录取信息管理
- 志愿填报表创建与编辑
- 志愿模拟录取结果预测
- 数据导入导出功能

## 运行界面
### 大学信息表界面图
![屏幕截图 2025-06-11 014657](https://github.com/user-attachments/assets/f7d0efbe-84d2-4077-9b27-347df8d08fd4)

### 高考志愿填报表界面图
![屏幕截图 2025-06-11 014715](https://github.com/user-attachments/assets/80c045b2-c32e-4c84-a3f6-d074ca5482d3)

### 志愿填报系统界面图
![屏幕截图 2025-06-11 013134](https://github.com/user-attachments/assets/76d8ec35-1869-4667-9663-3e739c77dc97)

### 使用界面
![屏幕截图 2025-06-11 030126](https://github.com/user-attachments/assets/3b034ca2-ee80-4766-b908-675df5347a98)

### 退档原则示意图
![屏幕截图 2025-06-11 121459](https://github.com/user-attachments/assets/f0c4d32a-d7d5-4f64-8124-c86556231952)

### 附件：福建高考报名表（与福建省教育考试院界面一致_官网有40个下图小表格，呈现为2x20排列）
![微信图片_20250611131701](https://github.com/user-attachments/assets/d84794ba-edb9-4b8a-b29c-d904263ea38a)




## 使用步骤（开发者版本）

### 1. 准备工作
确保已安装以下Python库：
```bash
pip install pandas openpyxl tk
```

### 2. 运行大学信息管理模块
```bash
python 大学信息表.py
```
- 功能：录入和管理各大学、专业组的录取信息
- 操作指南：
  1. 填写学校代号、名称等信息
  2. 添加专业组和专业信息
  3. 设置各专业的最低录取分数和录取人数
  4. 可通过"导出录取信息"按钮生成Excel文件

### 3. 运行志愿填报模块
```bash
python 模拟报考算法.py
```
- 功能：创建和编辑最多40个平行志愿
- 操作指南：
  1. 为每个志愿填写学校、专业组和专业信息
  2. 每个志愿最多可填写6个专业
  3. 设置是否服从调剂
  4. 点击"保存志愿"保存当前志愿
  5. 通过"导出所有志愿"生成Excel文件

### 4. 运行志愿模拟模块
```bash
python 模拟志愿填报.py
```
- 功能：模拟志愿录取过程
- 操作指南：
  1. 导入之前生成的"志愿填报表"和"大学录取信息表"
  2. 输入考生分数
  3. 点击"模拟填报"获取录取结果
  4. 系统将按照志愿顺序模拟录取过程

## 使用步骤（用户指南）

### 一、软件版本
1. **Python版**：需安装Python环境
2. **EXE版**：双击即可运行，无需安装Python

### 二、EXE版本使用说明

#### 1. 快速开始
1. 下载并解压软件包
2. 双击运行`志愿填报系统.exe`
3. 按照界面提示操作

#### 2. 功能模块
- **大学信息管理**：`大学信息表.exe`
- **志愿填报**：`模拟报考算法.exe`
- **志愿模拟**：`模拟志愿填报.exe`

#### 3. 数据文件说明
- 所有数据文件默认保存在`用户数据`文件夹中
- 支持导入/导出Excel文件(.xlsx格式)

### 三、详细使用步骤

#### 1. 大学信息录入
1. 运行`大学信息表.exe`
2. 点击"添加信息"填写：
   - 学校基本信息
   - 专业组信息
   - 各专业录取分数
3. 点击"导出录取信息"保存数据

#### 2. 志愿填报
1. 运行`模拟报考算法.exe`
2. 填写最多40个志愿：
   - 每个志愿填写1所学校和1-6个专业
   - 设置是否服从调剂
3. 点击"保存志愿"保存当前志愿
4. 点击"导出所有志愿"生成Excel文件

#### 3. 志愿模拟
1. 运行`模拟志愿填报.exe`
2. 导入"志愿填报表"和"大学录取信息表"
3. 输入考生分数
4. 点击"模拟填报"查看结果

### 四、数据格式要求
（同Python版本要求）

### 五、常见问题解答

#### EXE版本特有问题
Q: 运行时提示缺少DLL文件？
A: 请安装Visual C++ Redistributable运行库

Q: 软件无法启动？
A: 请右键选择"以管理员身份运行"

Q: 数据文件保存位置？
A: 默认保存在软件同级目录的"用户数据"文件夹中

### 六、注意事项
1. **兼容性**：EXE版本仅支持Windows 7/10/11系统
2. **安全提示**：从正规渠道下载软件，防止病毒
3. **数据备份**：定期备份数据文件
4. **更新说明**：新版发布时会提供升级包

### 七、联系支持
如遇问题请联系：TsangHaotian@hotmail.com
#### 微信联系：awa19490801


## 数据格式要求

### 大学录取信息表
必须包含以下列：
- 学校代号
- 学校名称
- 专业组代号
- 专业组名称
- 专业代号
- 专业名称
- 最低录取分
- 录取人数

### 志愿填报表
必须包含以下列：
- 志愿序号
- 学校代号
- 学校名称
- 专业组代号
- 专业组名称
- 专业代号1-6
- 专业名称1-6
- 是否服从
- 专业调剂

## 注意事项

1. **志愿顺序重要性**：系统严格按照志愿顺序进行模拟录取，请将最想去的学校/专业放在前面

2. **分数比对规则**：
   - 先比较是否达到专业组最低分数线
   - 再依次比较所填专业是否达到录取线
   - 最后考虑是否服从调剂

3. **调剂规则**：
   - 如果服从调剂，系统会自动调剂到该专业组内分数线最低且有名额的专业
   - 不服从调剂则直接退档，不再考虑后续志愿

4. **数据备份**：建议定期导出数据备份，防止意外丢失

5. **系统限制**：
   - 目前仅支持Excel格式(.xlsx)数据导入导出
   - 每个志愿最多填写6个专业
   - 模拟结果仅供参考，实际录取以考试院为准

## 常见问题

Q: 为什么导入Excel文件失败？
A: 请检查文件格式是否为.xlsx，且包含所有必需的列

Q: 模拟结果与实际录取不符？
A: 请检查大学录取信息表中的数据是否准确，特别是最低录取分数

Q: 如何增加更多志愿？
A: 目前系统设计为40个平行志愿，如需更多请修改代码中的相关限制

## 开发者说明
如需二次开发，请参考各模块的代码注释，主要数据结构如下：
- 大学信息使用嵌套字典存储
- 志愿信息使用DataFrame处理
- 模拟算法基于福建高考平行志愿规则实现

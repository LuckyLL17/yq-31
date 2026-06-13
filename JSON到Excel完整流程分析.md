# JSON数据从加载到Excel导出完整流程分析

## 一、整体架构概览

本项目是一个功能完善的JSON数据导出工具，采用**配置驱动**的模块化架构，支持将JSON数据导出为Excel、CSV、TSV、HTML、Markdown、JSON、PDF等多种格式。

```
┌───────────────────────────────────────────────────────────────────┐
│                      主入口：json_to_excel.py                     │
│  ┌────────┐  ┌────────────┐  ┌────────────┐  ┌────────────────┐ │
│  │ CLI参数│  │ 配置管理器  │  │ 数据加载器 │  │ 多格式导出器   │ │
│  └────────┘  └────────────┘  └────────────┘  └────────────────┘ │
│         │         │               │                │             │
│         ▼         ▼               ▼                ▼             │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │                  导出核心处理流程                        │   │
│  │  数据校验 → 计算列处理 → 数据写入 → 样式应用 → 透视表     │   │
│  └──────────────────────────────────────────────────────────┘   │
└───────────────────────────────────────────────────────────────────┘
```

---

## 二、核心模块文件清单

| 模块文件 | 核心职责 | 关键引用 |
|---------|---------|---------|
| [json_to_excel.py](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py) | 主入口、Excel导出核心逻辑、进度跟踪、数据透视表 | 核心处理文件 |
| [config_manager.py](file:///Volumes/ExMac/traeProject/全站1/yq-31/config_manager.py) | 配置加载、合并、验证、CRUD操作 | 配置驱动核心 |
| [multi_exporter.py](file:///Volumes/ExMac/traeProject/全站1/yq-31/multi_exporter.py) | 多格式导出调度、各格式具体实现 | 导出调度器 |
| [data_validator.py](file:///Volumes/ExMac/traeProject/全站1/yq-31/data_validator.py) | 数据校验规则定义、校验执行、错误处理 | 数据质量保障 |
| [computed_columns.py](file:///Volumes/ExMac/traeProject/全站1/yq-31/computed_columns.py) | 计算列公式解析、求值引擎 | 字段扩展 |
| [batch_processor.py](file:///Volumes/ExMac/traeProject/全站1/yq-31/batch_processor.py) | 批量文件处理、任务管理、断点续传 | 批量处理 |
| [style_template_manager.py](file:///Volumes/ExMac/traeProject/全站1/yq-31/style_template_manager.py) | 样式模板管理、条件格式匹配 | 样式系统 |
| [task_manager.py](file:///Volumes/ExMac/traeProject/全站1/yq-31/task_manager.py) | 批量任务持久化、进度跟踪 | 任务状态管理 |

---

## 三、JSON数据加载流程详解

### 3.1 加载入口函数

**函数签名**：[load_json(file_path)](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L135-L154)

```python
def load_json(file_path):
    # 1. 文件存在性检查
    # 2. UTF-8编码读取JSON文件
    # 3. 智能数据提取：支持多种包裹格式
    #    - 直接list → 原样返回
    #    - dict含data/items/records字段 → 提取对应list
    #    - 单个dict → 包装为list
    # 4. 非list非dict → 包装为list
    return data
```

### 3.2 数据扁平化处理

**函数签名**：[flatten_dict(d, parent_key="", sep=".")](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L157-L170)

递归遍历嵌套字典，使用点分隔符将嵌套key展平为单层key：
- `{"user": {"name": "张三", "age": 25}}` → `{"user.name": "张三", "user.age": 25}`
- 遇到list类型值 → 序列化为JSON字符串

### 3.3 字段自动检测

**函数签名**：[auto_detect_headers(data)](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L173-L187)

遍历所有数据项，收集所有出现过的key：
- 对每个数据项调用 `flatten_dict()` 获取扁平key
- 使用set去重后排序
- 自动生成label（将`.`替换为` / `，`_`替换为空格，首字母大写）
- 自动计算列宽（基于key长度）

### 3.4 表头合并策略

**函数签名**：[merge_headers(config_headers, auto_headers, config)](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L190-L199)

- 无配置表头 → 全部使用自动检测的表头
- 有配置表头 + `auto_detect_headers=True` → 配置表头 + 自动检测的额外字段
- 有配置表头 + `auto_detect_headers=False` → 仅使用配置表头

---

## 四、配置驱动机制详解

### 4.1 默认配置结构

**定义位置**：[DEFAULT_CONFIG](file:///Volumes/ExMac/traeProject/全站1/yq-31/config_manager.py#L6-L117)

配置采用分层结构，核心配置项包括：

| 配置层级 | 主要内容 |
|---------|---------|
| 基础配置 | json_file_path、export_format、各格式输出路径 |
| 表头配置 | default_headers（key/label/width）、auto_detect_headers |
| 样式配置 | header_style、data_style、conditional_format_rules |
| 校验配置 | validation_rules、validation_on_fail_default |
| 高级功能 | split_config、computed_columns、pivot_config |
| 格式专属 | csv_config、tsv_config、html_config等 |

### 4.2 配置加载与合并流程

**函数签名**：[load_config(config_path=None)](file:///Volumes/ExMac/traeProject/全站1/yq-31/config_manager.py#L126-L138)

```
配置文件不存在 → 返回默认配置副本
          ↓
配置文件存在 → 读取JSON → 调用merge_config()深度合并
          ↓
   合并策略：递归合并dict，非dict值直接覆盖
```

**函数签名**：[merge_config(base_config, override_config)](file:///Volumes/ExMac/traeProject/全站1/yq-31/config_manager.py#L154-L161)

### 4.3 配置验证机制

**函数签名**：[validate_config(config)](file:///Volumes/ExMac/traeProject/全站1/yq-31/config_manager.py#L164-L308)

验证项包括：
- 必填字段检查（json_file_path、输出路径）
- 导出格式有效性
- 表头配置格式（每个header必须有key）
- 校验规则格式（rule_type、on_fail有效性）
- 拆分配置格式（split_field、split_rule有效性）
- 条件格式规则验证
- 计算列配置验证
- 透视表配置验证

### 4.4 CLI参数覆盖机制

**函数签名**：[apply_cli_overrides(config, args)](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L1029-L1042)

命令行参数优先级 > 配置文件参数：
- `--input` → 覆盖 `json_file_path`
- `--format` → 覆盖 `export_format`
- `--output` → 根据格式覆盖对应输出路径

---

## 五、数据校验流程详解

### 5.1 校验规则体系

**类定义**：[ValidationRule](file:///Volumes/ExMac/traeProject/全站1/yq-31/data_validator.py#L91-L142)

支持的校验类型：
- `not_null`：非空校验
- `format`：格式校验（email/phone/url/date/number/id_card）
- `range`：数值范围校验（min/max）
- `regex`：正则表达式校验

失败处理策略（on_fail）：
- `mark`：标记该行，继续导出（高亮显示 + 批注）
- `skip`：跳过该行，不导出
- `abort`：中止整个导出流程

### 5.2 校验执行流程

**函数签名**：[validate_data(data, rules, headers=None)](file:///Volumes/ExMac/traeProject/全站1/yq-31/data_validator.py#L298-L324)

```
遍历每条数据
    ↓
flatten_dict()展平数据
    ↓
遍历每条校验规则
    ↓
提取字段值 → validate_value() → 失败则创建ValidationError
    ↓
根据on_fail策略更新ValidationResult
    ↓
遇到abort策略立即终止
```

### 5.3 校验结果应用

**函数签名**：[apply_validation_to_export(data, result, headers)](file:///Volumes/ExMac/traeProject/全站1/yq-31/data_validator.py#L327-L343)

- `aborted` → 返回None，导出取消
- `skipped_rows` → 过滤掉这些行
- `marked_rows` → 保留这些行，在Excel中应用标记样式

### 5.4 校验标记在Excel中的应用

**函数签名**：[_apply_validation_marks(ws, validation_result, headers, original_indices)](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L825-L861)

- 整行应用：红色加粗字体 + 黄色背景
- 特定字段：浅红色背景标记出错的单元格
- 添加批注：在A列添加包含所有错误信息的批注

---

## 六、计算列处理流程

### 6.1 计算列类型体系

**函数签名**：[evaluate_computed_column(col_config, item, extract_value_func, now=None)](file:///Volumes/ExMac/traeProject/全站1/yq-31/computed_columns.py#L303-L324)

支持的公式类型：

| 类型 | 功能说明 | 关键参数 |
|-----|---------|---------|
| `arithmetic` | 算术运算 | formula、referenced_fields |
| `date_diff` | 日期差值 | start_field、end_field、unit |
| `date_add` | 日期加减 | base_field、value、unit |
| `concat` | 字符串拼接 | parts、separator |
| `conditional` | 条件表达式 | condition、true_value、false_value |
| `round` | 四舍五入/取整 | field、precision、method |

### 6.2 计算列执行流程

**函数签名**：[apply_computed_columns(data, headers, config, extract_value_func)](file:///Volumes/ExMac/traeProject/全站1/yq-31/computed_columns.py#L327-L356)

```
1. 扩展headers：为每个启用的计算列添加header定义
2. 创建computed_cache字典：key=id(item)，value={计算列key: 计算值}
3. 遍历每条数据：
   - 提取所有referenced_fields的值
   - 根据formula_type调用对应求值函数
   - 结果存入computed_cache
4. 返回 (computed_cache, new_headers)
```

### 6.3 计算值的使用

在数据写入时，优先从computed_cache读取：

```python
if computed_cache and h["key"] in computed_cache.get(id(item), {}):
    val = computed_cache[id(item)][h["key"]]
else:
    val = extract_value(item, h["key"])
```

---

## 七、Excel导出主流程

### 7.1 导出入口函数

**函数签名**：[export_to_excel(data, headers, config)](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L703-L822)

完整执行流程：

```
┌─────────────────────────────────────────────────────────────┐
│ 步骤1：从配置生成校验规则 rules_from_config()               │
└─────────────────────────────┬───────────────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 步骤2：数据校验 validate_data()                             │
│  - 校验失败abort → 直接返回None                             │
│  - 校验失败skip → 过滤数据 apply_validation_to_export()     │
└─────────────────────────────┬───────────────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 步骤3：计算列处理 apply_computed_columns()                  │
│  - 生成computed_cache                                       │
│  - 扩展headers                                              │
└─────────────────────────────┬───────────────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 步骤4：判断是否拆分导出                                     │
│  - split_config.enabled=True → export_to_excel_with_split()│
│  - 否则 → 继续单sheet导出                                   │
└─────────────────────────────┬───────────────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 步骤5：创建Workbook + Worksheet                             │
└─────────────────────────────┬───────────────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 步骤6：写入表头 + 数据行（带进度跟踪）                       │
│  - ws.append(header_labels)                                 │
│  - 遍历数据行，extract_value()提取值                        │
│  - ProgressTracker.update()更新进度                         │
└─────────────────────────────┬───────────────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 步骤7：样式应用                                             │
│  - set_column_widths() 设置列宽                             │
│  - apply_header_style() 表头样式                            │
│  - apply_data_style() 数据行样式（含隔行变色）               │
│  - apply_conditional_formatting() 条件格式                  │
└─────────────────────────────┬───────────────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 步骤8：校验标记 + 冻结窗格 + 透视表                          │
│  - _apply_validation_marks() 校验错误标记                   │
│  - ws.freeze_panes = "A2" 冻结首行                          │
│  - add_pivot_table_to_workbook() 透视表                     │
└─────────────────────────────┬───────────────────────────────┘
                              ▼
┌─────────────────────────────────────────────────────────────┐
│ 步骤9：保存文件 wb.save(output_path)                        │
└─────────────────────────────────────────────────────────────┘
```

### 7.2 拆分导出模式

**函数签名**：[export_to_excel_with_split(data, headers, config, ...)](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L592-L700)

拆分规则（split_rule）：
- `by_value`：按字段值分组（每个唯一值一个sheet）
- `by_range`：按数值区间分组（配置range_groups）
- `by_custom`：按自定义规则分组（配置custom_rules，支持values/condition/min-max）

拆分流程：
1. `split_data_by_field()` 将数据分组
2. 可选创建"全部数据"汇总sheet
3. 为每个分组创建独立sheet，调用`_write_sheet_data()`
4. 自动处理sheet名称冲突（添加数字后缀）

### 7.3 数据写入核心函数

**函数签名**：[_write_sheet_data(ws, data, headers, config, ...)](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L551-L589)

这是一个通用的sheet数据写入函数，被单sheet导出和拆分导出共同调用，负责：
- 写入表头和数据行
- 设置列宽
- 应用表头样式、数据样式、条件格式
- 应用校验标记
- 设置冻结窗格

---

## 八、多格式导出机制

### 8.1 导出调度器

**函数签名**：[export_data(data, headers, config, fmt=None)](file:///Volumes/ExMac/traeProject/全站1/yq-31/multi_exporter.py#L539-L568)

```
1. 处理计算列（所有格式共用）
2. 根据fmt分发到具体导出函数：
   - excel → json_to_excel.export_to_excel()
   - csv → export_to_csv()
   - tsv → export_to_tsv()
   - html → export_to_html()
   - markdown → export_to_markdown()
   - json → export_to_json()
   - pdf → export_to_pdf()
```

### 8.2 格式专属配置

每种格式有独立的配置节点：

| 格式 | 配置节点 | 关键参数 |
|-----|---------|---------|
| CSV | csv_config | encoding、delimiter、quoting |
| TSV | tsv_config | encoding、include_header |
| HTML | html_config | title、style、custom_css |
| Markdown | markdown_config | title、max_col_width |
| JSON | json_config | indent、ensure_ascii、include_labels |
| PDF | pdf_config | page_size、orientation、font_size |

### 8.3 行数据准备

**函数签名**：[_prepare_rows(data, headers, computed_cache=None, ...)](file:///Volumes/ExMac/traeProject/全站1/yq-31/multi_exporter.py#L60-L83)

所有非Excel格式共用此函数准备二维数组数据，统一处理：
- 计算列值读取
- 原始字段值提取
- None值转空字符串
- 进度跟踪

---

## 九、批量处理流程

### 9.1 批量任务创建

**函数签名**：[start_batch_process(source_type, sources, output_dir, config, options)](file:///Volumes/ExMac/traeProject/全站1/yq-31/batch_processor.py#L203-L212)

支持的源类型：
- `directory`：递归扫描目录下所有JSON文件
- `files`：指定文件列表

任务创建内容：
- 生成唯一batch_id
- 扫描所有JSON文件，构建file_items列表
- 为每个文件生成输出路径（保持相对目录结构）
- 初始化文件状态为pending
- 持久化任务到磁盘

### 9.2 批量处理执行

**类定义**：[BatchProcessor](file:///Volumes/ExMac/traeProject/全站1/yq-31/batch_processor.py#L29-L200)

```
初始化信号处理器（支持Ctrl+C安全停止）
    ↓
设置任务状态为running
    ↓
获取待处理文件列表 get_pending_files()
    ↓
逐个处理文件 process_file():
    1. 更新文件状态为running
    2. load_json() 加载数据
    3. auto_detect_headers() 检测字段
    4. merge_headers() 合并表头
    5. export_data() 执行导出
    6. 更新状态为completed/failed/skipped
    ↓
更新批量任务状态（completed/failed/paused）
```

### 9.3 断点续传

**函数签名**：[resume_batch_process(batch_id)](file:///Volumes/ExMac/traeProject/全站1/yq-31/batch_processor.py#L215-L224)

- 从磁盘加载指定batch_id的任务
- 获取pending状态的文件（跳过已完成/失败的）
- 继续执行处理
- 支持`--batch-resume`命令行参数恢复

### 9.4 文件状态流转

```
pending → running → completed
                    → failed
                    → skipped（文件已存在且skip_existing=True）
```

---

## 十、数据透视表功能

### 10.1 透视表构建

**函数签名**：[build_pivot_table(data, pivot_config, headers)](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L1553-L1604)

透视表三维结构：
- 行字段（row_fields）：Y轴分组维度
- 列字段（column_fields）：X轴分组维度
- 值字段（value_fields）：聚合计算的度量

支持的聚合函数：
sum、count、count_num、average、max、min、product、stddev、stddevp、var、varp

### 10.2 透视表Sheet创建

**函数签名**：[create_pivot_sheet(wb, pivot_result, pivot_config, headers)](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L1739-L2012)

负责：
- 多级表头合并（列字段 + 值字段）
- 行标签写入
- 数据单元格计算与填充
- 行/列总计计算
- 样式应用（表头、隔行变色、边框）
- 冻结窗格设置

---

## 十一、函数间数据传递关系图

### 11.1 核心调用链（Excel导出）

```
main()
  │
  ├─► load_config() ◄─── DEFAULT_CONFIG
  │     │
  │     └─► merge_config()
  │
  ├─► validate_config()
  │
  └─► run_with_config()
        │
        ├─► load_json() ────────────────► data (List[dict])
        │
        ├─► auto_detect_headers(data) ──► auto_headers
        │
        ├─► merge_headers(config_headers, auto_headers) ──► headers
        │
        └─► export_data(data, headers, config, fmt="excel")
              │
              ├─► apply_computed_columns() ──► computed_cache, headers'
              │
              └─► export_to_excel(data, headers', config)
                    │
                    ├─► rules_from_config() ──► rules
                    │
                    ├─► validate_data(data, rules) ──► validation_result
                    │
                    ├─► apply_validation_to_export() ──► valid_data
                    │
                    ├─► (split分支) export_to_excel_with_split()
                    │     │
                    │     ├─► split_data_by_field() ──► groups
                    │     │
                    │     └─► _write_sheet_data() （每个group调用一次）
                    │
                    └─► (直接写入) 创建wb+ws
                          │
                          ├─► 写入表头和数据行（extract_value提取值）
                          │
                          ├─► set_column_widths()
                          │
                          ├─► apply_header_style()
                          │
                          ├─► apply_data_style()
                          │
                          ├─► apply_conditional_formatting()
                          │
                          ├─► _apply_validation_marks()
                          │
                          └─► add_pivot_table_to_workbook()
                                │
                                ├─► build_pivot_table() ──► pivot_result
                                │
                                └─► create_pivot_sheet()
```

### 11.2 关键数据对象传递

| 数据对象 | 类型 | 传递路径 | 核心内容 |
|---------|------|---------|---------|
| `data` | `List[dict]` | load_json → auto_detect_headers → validate_data → apply_validation_to_export → export_to_excel | 原始JSON数据列表 |
| `headers` | `List[dict]` | auto_detect_headers → merge_headers → apply_computed_columns → _write_sheet_data | 字段定义列表，每个元素含key/label/width |
| `config` | `dict` | load_config → apply_cli_overrides → validate_config → run_with_config → 所有导出函数 | 完整配置字典 |
| `validation_result` | `ValidationResult` | validate_data → apply_validation_to_export → _apply_validation_marks | 校验结果，含errors/marked_rows/skipped_rows |
| `computed_cache` | `Dict[int, Dict]` | apply_computed_columns → _prepare_rows / _write_sheet_data | 计算列缓存，key=id(item)，value={计算列key: 值} |
| `original_indices` | `List[int]` | export_to_excel → export_to_excel_with_split → _write_sheet_data | 数据行在原始列表中的索引，用于校验标记映射 |

---

## 十二、配置驱动机制完整图示

```
┌───────────────────────────────────────────────────────────────────┐
│                        配置来源层级                                │
└───────────────────────────────────────────────────────────────────┘
                                  │
                                  ▼
┌───────────────────────────────────────────────────────────────────┐
│  1. DEFAULT_CONFIG (config_manager.py)                            │
│     内置默认值，所有字段的兜底配置                                 │
└───────────────────────────────────┬───────────────────────────────┘
                                    │ merge_config() 深度合并
                                    ▼
┌───────────────────────────────────────────────────────────────────┐
│  2. config.json (用户配置文件)                                    │
│     用户自定义配置，覆盖默认值                                     │
└───────────────────────────────────┬───────────────────────────────┘
                                    │ apply_cli_overrides()
                                    ▼
┌───────────────────────────────────────────────────────────────────┐
│  3. CLI命令行参数 (--input/--output/--format等)                   │
│     最高优先级，临时覆盖配置                                       │
└───────────────────────────────────┬───────────────────────────────┘
                                    │
                                    ▼
┌───────────────────────────────────────────────────────────────────┐
│  4. 运行时动态配置 (向导模式/预览筛选)                             │
│     _filtered_data、交互式配置修改                                 │
└───────────────────────────────────┬───────────────────────────────┘
                                    │ validate_config()
                                    ▼
                          最终生效的配置对象
```

---

## 十三、关键设计模式与技术亮点

### 13.1 策略模式

- 多格式导出：根据`export_format`选择不同的导出策略
- 校验规则：4种校验类型实现为不同的验证函数
- 拆分规则：3种拆分策略（by_value/by_range/by_custom）
- 聚合函数：12种聚合函数在透视表中动态调用

### 13.2 模板方法模式

- `_write_sheet_data()` 作为模板方法，被单sheet和拆分导出复用
- `_prepare_rows()` 作为行数据准备模板，被所有非Excel格式复用

### 13.3 安全设计

- `eval()` 表达式求值使用沙箱环境（`__builtins__: {}` + 白名单函数）
- 条件表达式求值同样限制可用函数
- 文件路径安全校验，防止批量处理时路径穿越

### 13.4 进度跟踪

**类定义**：[ProgressTracker](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L14-L113)

- 实时进度条渲染
- 速度/已用时间/剩余时间估算
- 当前处理字段和行预览
- 防闪烁渲染（最小更新间隔）

---

## 十四、主入口函数完整执行路径

**函数签名**：[main()](file:///Volumes/ExMac/traeProject/全站1/yq-31/json_to_excel.py#L1289-L1426)

```
1. 解析命令行参数 parse_args()
    │
    ├─► --list-formats → 列出支持格式并退出
    │
    ├─► 批量操作处理 _handle_batch_operations()
    │   ├─► --batch → 启动批量处理向导
    │   ├─► --batch-list → 列出批量任务
    │   ├─► --batch-resume → 恢复批量任务
    │   ├─► --batch-report → 生成批量报告
    │   └─► --batch-dir/--batch-files → 直接批量处理
    │
    ├─► --validate → 启动校验规则配置向导
    │
    ├─► --validate-only → 仅执行校验不导出
    │
    ├─► --computed-columns → 启动计算列配置向导
    │
    ├─► --wizard → 启动交互式配置向导
    │   └─► run_wizard() → 保存配置 → 可选执行导出
    │
    ├─► 模板操作 handle_template_operations()
    │   ├─► --list-templates → 列出样式模板
    │   ├─► --template-manager → 启动模板管理器
    │   ├─► --apply-template → 应用指定模板
    │   └─► --save-as-template → 保存当前样式为模板
    │
    ├─► --preview → 数据预览与筛选模式
    │   └─► start_preview_mode() → 可选导出筛选后数据
    │
    └─► 正常导出流程
          ├─► load_config()
          ├─► apply_cli_overrides()
          ├─► validate_config()
          ├─► 可选：数据预览
          └─► run_with_config()
```

---

## 十五、总结

### 15.1 数据流总览

```
JSON文件 → load_json() → 原始数据List
              │
              ▼
         auto_detect_headers() → 自动检测字段
              │
              ▼
         merge_headers() → 最终字段定义
              │
              ▼
         validate_data() → 校验结果
              │
              ▼
         apply_computed_columns() → 计算列缓存 + 扩展字段
              │
              ▼
         export_to_excel() 或 其他格式导出函数
              │
              ▼
         输出文件（Excel/CSV/TSV/HTML/Markdown/JSON/PDF）
```

### 15.2 核心设计思想

1. **配置驱动**：所有行为通过配置控制，代码与逻辑分离
2. **模块化**：每个功能独立成模块，职责清晰，易于维护和扩展
3. **渐进式增强**：从基础导出 → 样式 → 校验 → 计算列 → 透视表 → 批量处理，层层递进
4. **用户友好**：提供交互式向导、进度跟踪、实时预览等体验优化
5. **容错设计**：断点续传、错误恢复、跳过策略等保障批量处理可靠性

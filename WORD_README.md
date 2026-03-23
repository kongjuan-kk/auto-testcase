# Word 文档读取指南

## auto-testcase skill 现已支持 Word 文档！

### 支持的文件格式

- ✅ `.txt` - 纯文本文件
- ✅ `.md` - Markdown 文件
- ✅ `.docx` - Word 文档（2007及更高版本）
- ❌ `.doc` - 旧版 Word 文档（不支持，请先转换为 .docx）

### 使用方法

#### 方法1：直接提供 Word 文档路径

```
请为这个需求文档生成测试用例：C:\path\to\requirement.docx
```

auto-testcase skill 会自动：
1. 使用 `scripts/read_word.py` 读取 Word 文档
2. 提取文本内容（包括段落和表格）
3. 生成测试用例

#### 方法2：手动转换后提供

如果自动读取失败，可以手动转换：

```bash
# 使用 skill 提供的脚本
python D:\Skills\auto-testcase\scripts\read_word.py "C:\path\to\requirement.docx" > output.txt

# 然后提供 output.txt 的内容
```

### Word 文档内容要求

为确保测试用例生成质量，Word 文档应包含：

1. **功能描述**：清晰的功能说明
2. **输入参数**：字段名称、类型、限制条件
3. **业务规则**：验证规则、约束条件
4. **输出结果**：预期结果描述
5. **异常场景**：可能的错误情况

#### 推荐的文档结构

```
1. 功能概述
   - 功能名称
   - 功能描述

2. 功能点
   2.1 搜索功能
       - 搜索条件1：字段名称、类型、长度限制
       - 搜索条件2：字段名称、类型、长度限制

3. 业务规则
   - 规则1
   - 规则2

4. 界面说明
   - 表格/字段列表
```

### 表格内容提取

脚本会自动提取 Word 文档中的表格内容，例如：

| 字段名称 | 字段类型 | 长度限制 | 必填 |
|---------|---------|---------|------|
| 门店编码 | 文本 | 20字符 | 是 |

会被转换为：
```
字段名称 | 字段类型 | 长度限制 | 必填
门店编码 | 文本 | 20字符 | 是
```

### 故障排查

#### 问题1：python-docx 未安装

**错误信息**：
```
ERROR: python-docx is not installed.
Install it with: pip install python-docx
```

**解决方案**：
```bash
pip install python-docx
```

#### 问题2：文档包含复杂格式

如果 Word 文档包含复杂格式（如图片、复杂表格），建议：
1. 将文档另存为 .txt 或 .md 格式
2. 复制文本内容直接粘贴给 skill

#### 问题3：编码问题

如果文档包含特殊字符出现乱码：
1. 确保文档保存为 UTF-8 编码
2. 或使用 .txt 格式提供需求

### 示例

#### 输入：Word 文档

```
请为以下文档生成测试用例：C:\Users\lengj\Desktop\需求文档.docx
```

#### 处理流程

1. Skill 检测到 .docx 文件
2. 自动调用 `read_word.py` 提取内容
3. 分析需求文档内容
4. 生成测试用例
5. 输出 Excel 文件

#### 输出：Excel 测试用例

```
✅ 测试用例生成成功！
文件路径：C:\Users\lengj\Desktop\测试用例_XXX功能_20260318.xlsx
用例数量：99 条
```

### 高级功能

#### 批量处理多个 Word 文档

如果有多个需求文档：

```bash
# 创建批处理脚本
for file in *.docx; do
    python D:\Skills\auto-testcase\scripts\read_word.py "$file" > "${file%.docx}.txt"
done

# 然后逐个处理生成的 .txt 文件
```

#### 与其他工具集成

可以与以下工具配合使用：
- **Pandoc**：万能文档转换器
- **LibreOffice**：打开旧版 .doc 文档并另存为 .docx
- **Markdown**：将 Word 转换为 Markdown 后处理

### 反馈与建议

如有问题或改进建议，请联系技能维护者。

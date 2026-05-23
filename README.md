# 教案合规性检测智能体

基于LangChain + DeepSeek + Qwen-VL的教案合规性自动检测系统


##  系统界面
<img width="1919" height="921" alt="QQ20260523-224239" src="https://github.com/user-attachments/assets/42a21d84-c345-478f-b4d1-673adfe305f1" />

## 🌟 功能特性

- ✅ **智能解析**：自动解析Markdown格式的教案文档，提取文本段落和图片
- ✅ **文本检测**：使用DeepSeek大语言模型检测文本内容的合规性
- ✅ **图片检测**：使用Qwen-VL多模态模型检测图片及上下文的合规性
- ✅ **多条规则**：基于19条预定义规则进行全面的合规性审查
- ✅ **详细报告**：生成包含总体得分、风险等级、违规详情的检测报告
- ✅ **精准定位**：准确定位违规段落的页码和行号
- ✅ **批量处理**：支持单文件检测和文件夹批量检测

## 📋 合规规则

系统基于以下19条规则进行检测：

1. **政治敏感内容检测** - 不得包含违反国家政治立场的内容
2. **色情低俗内容检测** - 不得包含不适合教学的内容
3. **暴力血腥内容检测** - 不得包含可能造成心理伤害的内容
4. **歧视性内容检测** - 不得包含种族、性别等歧视性内容
5. **违法违规内容检测** - 不得包含教唆犯罪等内容
6. **虚假信息检测** - 不得包含伪科学、迷信等内容
7. **版权侵权检测** - 应注明来源，避免侵权
8. **不当商业宣传检测** - 不得包含商业广告
9. **年龄适宜性检测** - 内容应符合学生年龄段
10. **语言规范性检测** - 使用规范的语言文字
....
## 🚀 快速开始

### 1. 安装依赖

**重要提示**：本项目需要 LangChain 1.0+ 版本

```bash
# 方法1：一键安装（Windows）
install.bat

# 方法2：手动安装（所有平台）
pip install -r requirements.txt

# 如果遇到问题，请查看 QUICK_START.md
```

### 2. 配置API密钥

复制`.env.example`为`.env`，并填写你的API密钥：

```bash
cp .env.example .env
```

编辑`.env`文件：

```
# DeepSeek API配置
DEEPSEEK_API_KEY=sk-5e1bd86fbeb04b80ab0a75d57ad
DEEPSEEK_BASE_URL=https://api.deepseek.com

# Qwen-VL API配置（多模态大模型）
# 注意：该服务不需要API密钥
QWEN_VL_BASE_URL=http://xn-d.suanjiayun.com:57915/v1
QWEN_VL_MODEL=Qwen/Qwen3-VL-235B-A22B-Instruct-FP8
```

### 3. 准备教案文档

确保你的Markdown教案文档格式如下：

```
your_lesson.md        # 教案文档
image/                # 图片文件夹
  ├── pic1.png
  ├── pic2.jpg
  └── ...
```

### 4. 运行检测

#### 检测单个文档

```bash
python agent.py --file your_lesson.md
```

#### 批量检测文件夹

```bash
python agent.py --folder ./lessons
```

#### 自定义图片文件夹

```bash
python agent.py --file your_lesson.md --image-folder images
```

#### 自定义报告输出目录

```bash
python agent.py --file your_lesson.md --output my_reports
```

## 📊 检测报告

检测完成后，系统会生成两种格式的报告：

### JSON报告 (`*_report.json`)

包含结构化的检测数据，便于程序化处理：

```json
{
  "document_name": "lesson.md",
  "检测时间": "2024-01-01 12:00:00",
  "总体得分": 85.5,
  "风险等级": "中风险",
  "检测统计": {...},
  "违规段落详情": [...]
}
```

### Markdown报告 (`*_report.md`)

人类可读的格式化报告，包含：

- **基本信息**：文档名称、检测时间、总体得分、风险等级
- **检测统计**：段落数、合规率等统计数据
- **违规详情**：每处违规的详细信息
  - 位置（页码、段落、行号）
  - 严重程度
  - 内容预览
  - 违反的规则
  - 修改建议
  - 法律依据
- **总体建议**：针对整个文档的改进建议

## 📁 项目结构

```
.
├── agent.py                    # 主程序
├── markdown_parser.py          # Markdown解析器
├── compliance_checker.py       # 合规性检测器
├── report_generator.py         # 报告生成器
├── compliance_rules.json       # 合规规则配置
├── requirements.txt            # Python依赖
├── .env.example               # 环境变量示例
├── README.md                  # 项目说明
└── examples/                  # 示例文件
    └── sample_lesson.md       # 示例教案
```

## 🔧 高级用法

### 自定义合规规则

编辑`compliance_rules.json`文件，可以添加、修改或删除规则：

```json
{
  "rules": [
    {
      "id": 11,
      "name": "自定义规则名称",
      "description": "规则描述",
      "severity": "high",
      "keywords": ["关键词1", "关键词2"],
      "category": "分类名称"
    }
  ]
}
```

### 作为Python模块使用

```python
from agent import ComplianceAgent

# 初始化智能体
agent = ComplianceAgent(
    deepseek_api_key="your_key",
    deepseek_base_url="https://api.deepseek.com",
    qwen_vl_api_key="your_key",
    qwen_vl_base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"
)

# 检测文档
violations = agent.check_document("lesson.md")

# 生成报告
agent.generate_report("lesson.md", violations)
```

## 🔍 工作原理

1. **文档解析**：使用`MarkdownParser`解析Markdown文档，提取文本段落和图片信息
2. **文本检测**：将文本段落发送给DeepSeek，根据10条规则判断是否合规
3. **图片检测**：将图片及上下文发送给Qwen-VL多模态模型进行检测
4. **结果汇总**：收集所有检测结果，计算总体得分和风险等级
5. **报告生成**：生成详细的JSON和Markdown格式报告

## ⚠️ 注意事项

1. **API费用**：使用DeepSeek和Qwen-VL API会产生费用，请注意控制使用量
2. **图片格式**：支持常见图片格式（jpg, png, gif等）
3. **文档大小**：单个文档不建议超过1000段落，否则检测时间会较长
4. **准确性**：AI检测结果仅供参考，重要文档请人工复核

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个项目！

## 📄 许可证

MIT License

## 💡 常见问题

### Q: 如何获取API密钥？

**DeepSeek**: 访问 https://platform.deepseek.com/ 注册并获取API密钥  
**Qwen-VL**: 访问阿里云灵积平台 https://dashscope.aliyun.com/ 获取API密钥

### Q: 检测速度慢怎么办？

- 减少文档段落数量
- 使用更快的API节点
- 批量检测时可以并行处理（需修改代码）

### Q: 可以检测其他格式的文档吗？

目前仅支持Markdown格式。如需支持其他格式（如Word、PDF），需要添加相应的解析器。

### Q: 检测结果不准确怎么办？

- 调整`compliance_rules.json`中的规则描述
- 修改提示词模板以提高准确性
- 使用更强大的模型

## 📮 联系方式

如有问题或建议，请提交GitHub Issue。

---

**声明**：本系统由AI驱动，检测结果仅供参考，最终解释权归相关监管部门所有。

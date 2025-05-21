# AzureAI
Azure AI 服务集成示例

这个仓库展示了如何使用Python集成Azure OpenAI服务。

## 功能特点

- 支持Azure OpenAI的聊天补全 (Chat Completion) API
- 支持文本补全 (Text Completion) API
- 支持嵌入向量 (Embeddings) API
- 自动配置和验证环境设置
- 提供命令行界面和交互模式

## 安装

1. 安装依赖包:

```bash
pip install python-dotenv requests
```

2. 配置Azure OpenAI:
   - 在Azure门户中创建Azure OpenAI服务资源
   - 创建模型部署
   - 复制API密钥和终端点URL

3. 配置环境变量:
   - 首次运行程序时会自动创建.env模板文件
   - 编辑.env文件，添加你的Azure OpenAI凭据

## 使用方法

### 测试配置

```bash
python main.py --mode test
```

### 聊天补全

```bash
python main.py --mode chat --input "你好，请介绍一下Azure AI服务"
```

### 文本补全

```bash
python main.py --mode completion --input "Azure OpenAI是"
```

### 嵌入向量

```bash
python main.py --mode embeddings --input "将这段文本转换为嵌入向量"
```

### 输出结果到文件

```bash
python main.py --mode chat --input "解释什么是机器学习" --output result.json
```

### 高级参数

设置温度参数 (0-1 之间):

```bash
python main.py --mode chat --input "创意写作" --temp 0.9
```

## 项目结构

- `main.py`: 主程序，包含Azure OpenAI集成代码
- `.env`: 环境变量配置文件
- `README.md`: 说明文档

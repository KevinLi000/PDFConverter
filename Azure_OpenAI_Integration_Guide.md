# Azure OpenAI 集成指南

## 解决 "No model specified in request" 错误

如果您在调用 Azure OpenAI 服务时遇到以下错误：

```
No model specified in request. Please provide a model name in the request body or as a x-ms-model-mesh-model-name header
```

这是因为在请求体中缺少 `model` 参数。以下是解决方法：

### 1. 修复请求格式

在您的 API 请求中，确保在请求数据中包含 `model` 参数。例如：

```python
data = {
    "messages": messages,
    "temperature": temperature,
    "max_tokens": max_tokens,
    "model": self.deployment  # 添加模型参数
}
```

### 2. 确认部署名称

确保您的 `.env` 文件中设置了正确的部署名称。部署名称应该与您在 Azure OpenAI 服务中创建的部署相匹配：

```
AZURE_OPENAI_DEPLOYMENT=gpt-35-turbo
```

### 3. 检查可用模型

您可以使用以下端点查看您账户中可用的模型：
```
https://your-resource-name.openai.azure.com/openai/models?api-version=2023-05-15
```

### 4. 修复后的代码

本项目中的 `main_new.py` 文件已经修复了这个问题。主要修复包括：

1. 在所有 API 请求中添加 `model` 参数
2. 修复了代码中的缩进问题
3. 添加了模型检查功能，验证部署名称是否有效

## 运行说明

使用以下命令运行程序：

```bash
# 测试配置
python main_new.py --mode test

# 聊天模式
python main_new.py --mode chat --input "你好，我想了解一下Azure AI服务"

# 文本补全模式
python main_new.py --mode completion --input "Azure OpenAI 是一种"

# 嵌入向量模式
python main_new.py --mode embeddings --input "这是一段用于生成向量表示的文本"
```

## 环境变量设置

确保您的 `.env` 文件包含以下内容：

```
# Azure OpenAI 配置
AZURE_OPENAI_KEY=您的API密钥
AZURE_OPENAI_ENDPOINT=https://您的资源名称.openai.azure.com/
AZURE_OPENAI_API_VERSION=2023-05-15
AZURE_OPENAI_DEPLOYMENT=您的部署名称或模型ID
```

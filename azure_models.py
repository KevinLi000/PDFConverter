#!/usr/bin/env python
# filepath: c:\Users\kevin.li\OneDrive - GREEN DOT CORPORATION\Documents\GitHub\Azure\AzureAI\azure_models.py

"""
Azure OpenAI 模型信息工具
此工具帮助用户了解各 API 版本支持的 Azure OpenAI 模型
"""

import json
import requests
import os
from dotenv import load_dotenv
import argparse
from tabulate import tabulate

# API 版本与支持模型的映射关系
API_VERSION_MODELS = {
    "2025-05-15": {
        "GPT 模型": [
            {"名称": "gpt-4o", "描述": "最新的多模态模型，支持图像、音频和视频输入，上下文窗口 128K tokens"},
            {"名称": "gpt-4-turbo", "描述": "GPT-4 的优化版本，成本更低，速度更快，上下文窗口 128K tokens"},
            {"名称": "gpt-4-32k", "描述": "扩展上下文长度的 GPT-4 模型，上下文窗口 32K tokens"},
            {"名称": "gpt-4", "描述": "高级推理和指令跟随能力的模型，上下文窗口 8K tokens"},
            {"名称": "gpt-35-turbo-16k", "描述": "扩展上下文长度的 GPT-3.5 模型，上下文窗口 16K tokens"},
            {"名称": "gpt-35-turbo", "描述": "平衡性能和成本的模型，上下文窗口 4K tokens"}
        ],
        "嵌入模型": [
            {"名称": "text-embedding-ada-002", "描述": "通用文本嵌入模型，生成 1536 维向量"},
            {"名称": "text-embedding-3-large", "描述": "高级嵌入模型，生成 3072 维向量，提供更高质量的语义表示"},
            {"名称": "text-embedding-3-small", "描述": "平衡性能和效率的嵌入模型，生成 1536 维向量"}
        ],
        "图像模型": [
            {"名称": "dall-e-3", "描述": "最新的图像生成模型，支持高分辨率和复杂提示"},
            {"名称": "dall-e-2", "描述": "原始的 DALL-E 图像生成模型"}
        ],
        "语音模型": [
            {"名称": "whisper-multilingual", "描述": "支持多语言的语音识别模型"},
            {"名称": "speech-synthesis-1", "描述": "文本转语音合成模型，支持多种自然语音"}
        ]
    },
    "2024-05-01": {
        "GPT 模型": [
            {"名称": "gpt-4-turbo", "描述": "GPT-4 的优化版本，上下文窗口 128K tokens"},
            {"名称": "gpt-4-32k", "描述": "扩展上下文长度的 GPT-4 模型，上下文窗口 32K tokens"},
            {"名称": "gpt-4", "描述": "高级推理和指令跟随能力的模型，上下文窗口 8K tokens"},
            {"名称": "gpt-35-turbo-16k", "描述": "扩展上下文长度的 GPT-3.5 模型，上下文窗口 16K tokens"},
            {"名称": "gpt-35-turbo", "描述": "平衡性能和成本的模型，上下文窗口 4K tokens"}
        ],
        "嵌入模型": [
            {"名称": "text-embedding-ada-002", "描述": "通用文本嵌入模型，生成 1536 维向量"},
            {"名称": "text-embedding-3-small", "描述": "平衡性能和效率的嵌入模型，生成 1536 维向量"}
        ],
        "图像模型": [
            {"名称": "dall-e-3", "描述": "高级图像生成模型"},
            {"名称": "dall-e-2", "描述": "原始的 DALL-E 图像生成模型"}
        ],
        "语音模型": [
            {"名称": "whisper", "描述": "基础语音识别模型"}
        ]
    },
    "2023-05-15": {
        "GPT 模型": [
            {"名称": "gpt-4", "描述": "高级推理和指令跟随能力的模型，上下文窗口 8K tokens"},
            {"名称": "gpt-35-turbo", "描述": "平衡性能和成本的模型，上下文窗口 4K tokens"},
            {"名称": "text-davinci-003", "描述": "针对文本补全的老式模型"}
        ],
        "嵌入模型": [
            {"名称": "text-embedding-ada-002", "描述": "通用文本嵌入模型，生成 1536 维向量"}
        ],
        "图像模型": [
            {"名称": "dall-e-2", "描述": "图像生成模型"}
        ]
    }
}

def get_available_models(endpoint, api_key, api_version):
    """
    从 Azure OpenAI 服务获取可用的模型列表
    
    参数:
        endpoint (str): Azure OpenAI 终端点
        api_key (str): API 密钥
        api_version (str): API 版本
        
    返回:
        list: 可用模型列表
    """
    try:
        url = f"{endpoint}/openai/models?api-version={api_version}"
        headers = {
            "api-key": api_key,
            "Content-Type": "application/json"
        }
        
        print(f"正在从 Azure OpenAI 服务获取可用模型...")
        
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            result = response.json()
            models = result.get("data", [])
            return models
        else:
            print(f"获取模型失败 (状态码: {response.status_code})")
            print(f"错误: {response.text}")
            return []
    except Exception as e:
        print(f"获取模型时发生错误: {str(e)}")
        return []

def display_models_for_version(api_version):
    """
    显示指定 API 版本支持的模型
    
    参数:
        api_version (str): API 版本
    """
    if api_version in API_VERSION_MODELS:
        print(f"\n{api_version} 版本支持的模型:")
        print("=" * 50)
        
        models_by_category = API_VERSION_MODELS[api_version]
        
        for category, models in models_by_category.items():
            print(f"\n{category}:")
            table_data = [[m["名称"], m["描述"]] for m in models]
            print(tabulate(table_data, headers=["模型名称", "描述"], tablefmt="grid"))
    else:
        print(f"\n{api_version} 版本的模型信息不可用")
        print("请尝试以下已知的 API 版本:")
        for version in API_VERSION_MODELS.keys():
            print(f"  • {version}")

def display_actual_available_models(models):
    """
    显示实际可用的模型
    
    参数:
        models (list): 模型列表
    """
    if not models:
        print("\n未找到可用的模型")
        return
    
    print("\n您账户中实际可用的模型:")
    print("=" * 50)
    
    # 按模型类型分组
    models_by_type = {}
    for model in models:
        model_id = model.get("id", "未知")
        model_type = "GPT 模型"
        
        if "embedding" in model_id:
            model_type = "嵌入模型"
        elif "dall-e" in model_id:
            model_type = "图像模型"
        elif "whisper" in model_id or "speech" in model_id:
            model_type = "语音模型"
        
        if model_type not in models_by_type:
            models_by_type[model_type] = []
        
        models_by_type[model_type].append(model)
    
    # 显示每种类型的模型
    for model_type, type_models in models_by_type.items():
        print(f"\n{model_type}:")
        table_data = []
        for model in type_models:
            model_id = model.get("id", "未知")
            model_created = model.get("created", "未知")
            model_owner = model.get("owned_by", "未知")
            table_data.append([model_id, model_owner, model_created])
        
        print(tabulate(table_data, headers=["模型名称", "提供者", "创建时间"], tablefmt="grid"))

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="Azure OpenAI 模型信息工具")
    parser.add_argument("--version", default="2025-05-15", 
                      help="要查询的 API 版本 (默认: 2025-05-15)")
    parser.add_argument("--check-available", action="store_true", 
                      help="检查账户中实际可用的模型")
    parser.add_argument("--list-versions", action="store_true",
                      help="列出所有已知的 API 版本")
    
    args = parser.parse_args()
    
    # 列出所有已知的 API 版本
    if args.list_versions:
        print("\n已知的 API 版本:")
        for version in sorted(API_VERSION_MODELS.keys(), reverse=True):
            print(f"  • {version}")
        return
    
    # 显示指定 API 版本支持的模型
    display_models_for_version(args.version)
    
    # 检查实际可用的模型
    if args.check_available:
        # 加载环境变量
        load_dotenv()
        
        # 获取配置信息
        endpoint = os.environ.get("AZURE_OPENAI_ENDPOINT", "")
        api_key = os.environ.get("AZURE_OPENAI_KEY", "")
        api_version = os.environ.get("AZURE_OPENAI_API_VERSION", args.version)
        
        if not endpoint or not api_key:
            print("\n警告: 未设置 Azure OpenAI 凭据，无法检查实际可用的模型")
            print("请在 .env 文件中设置 AZURE_OPENAI_ENDPOINT 和 AZURE_OPENAI_KEY")
            return
        
        # 获取实际可用的模型
        available_models = get_available_models(endpoint, api_key, api_version)
        display_actual_available_models(available_models)
    else:
        print("\n提示: 使用 --check-available 参数可检查您账户中实际可用的模型")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"程序运行出错: {str(e)}")

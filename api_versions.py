#!/usr/bin/env python
# filepath: c:\Users\kevin.li\OneDrive - GREEN DOT CORPORATION\Documents\GitHub\Azure\AzureAI\api_versions.py

"""
Azure OpenAI API 版本工具
此工具帮助用户了解可用的 Azure OpenAI API 版本
"""

import json
import requests
import os
from dotenv import load_dotenv
import argparse

# 已知的 API 版本列表（截至 2025 年 5 月）
KNOWN_API_VERSIONS = [
    "2025-05-15",        # 最新稳定版
    "2024-07-01-preview", # 引入高级上下文处理和图像分析
    "2024-05-01",        # 稳定版，改进的函数调用
    "2023-12-01-preview", # GPT-4 流式处理和系统提示增强
    "2023-09-15-preview", # 批处理请求支持
    "2023-07-01-preview", # GPT-4V 视觉能力
    "2023-06-01-preview", # 多模态模型支持
    "2023-05-15",        # 基本聊天完成、嵌入和图像生成
    "2022-12-01"         # 早期稳定版
]

# API 版本功能映射
API_VERSION_FEATURES = {
    "2025-05-15": [
        "支持 GPT-4-32k 和最新模型",
        "RAG（检索增强生成）增强功能",
        "改进的 JSON 模式支持",
        "强化的系统提示和函数调用",
        "高级多模态处理：图像、音频和视频",
        "低延迟，高吞吐量"
    ],
    "2024-05-01": [
        "改进的函数调用和工具使用",
        "增强的流式处理功能",
        "系统级错误处理改进",
        "更好的并发请求支持"
    ],
    "2023-05-15": [
        "基础聊天完成 API",
        "文本嵌入功能",
        "基本函数调用",
        "DALL-E 图像生成（初始版本）"
    ]
}

def get_api_versions(endpoint, api_key, current_api_version):
    """
    尝试从 Azure OpenAI 服务获取可用的 API 版本信息
    
    参数:
        endpoint (str): Azure OpenAI 终端点
        api_key (str): API 密钥
        current_api_version (str): 当前使用的 API 版本
        
    返回:
        list: API 版本列表
    """
    try:
        # 尝试使用元数据端点获取信息（这是一个假设的端点，可能需要根据实际情况调整）
        url = f"{endpoint}/openai/info?api-version={current_api_version}"
        headers = {
            "api-key": api_key,
            "Content-Type": "application/json"
        }
        
        print(f"尝试获取API版本信息...")
        
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            result = response.json()
            print("成功获取API信息")
            
            if "api_versions" in result:
                versions = result["api_versions"]
                return versions
            else:
                print("响应中未找到API版本信息，使用已知版本列表")
        else:
            print(f"无法直接获取API版本信息 (状态码: {response.status_code})")
    except Exception as e:
        print(f"获取API版本信息时发生错误: {str(e)}")
    
    # 返回已知的版本列表作为备选
    print("提供已知的API版本列表...")
    return KNOWN_API_VERSIONS

def display_version_features(version):
    """
    显示特定 API 版本的功能特性
    
    参数:
        version (str): API 版本
    """
    if version in API_VERSION_FEATURES:
        print(f"\n{version} 版本主要特性:")
        for feature in API_VERSION_FEATURES[version]:
            print(f"  • {feature}")
    else:
        print(f"\n{version} 版本的详细特性信息不可用")
        if version.endswith("-preview"):
            print("  • 这是一个预览版本，可能包含实验性功能")
            print("  • 预览版本通常不建议用于生产环境")

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="Azure OpenAI API 版本信息工具")
    parser.add_argument("--detail", action="store_true", help="显示详细功能信息")
    parser.add_argument("--version", help="指定要查看详情的API版本")
    
    args = parser.parse_args()
    
    # 加载环境变量
    load_dotenv()
    
    # 获取配置信息
    endpoint = os.environ.get("AZURE_OPENAI_ENDPOINT", "")
    api_key = os.environ.get("AZURE_OPENAI_KEY", "")
    current_api_version = os.environ.get("AZURE_OPENAI_API_VERSION", "2023-05-15")
    
    # 显示当前配置
    print("\nAzure OpenAI API 版本信息")
    print("=======================")
    print(f"当前配置的 API 版本: {current_api_version}")
    
    # 如果指定了特定版本，显示其详情
    if args.version:
        display_version_features(args.version)
        return
    
    # 获取并显示可用版本
    if endpoint and api_key:
        versions = get_api_versions(endpoint, api_key, current_api_version)
    else:
        print("警告: 未设置 Azure OpenAI 凭据，使用已知版本列表")
        versions = KNOWN_API_VERSIONS
    
    print("\n可用的 API 版本:")
    stable_versions = [v for v in versions if not v.endswith("-preview")]
    preview_versions = [v for v in versions if v.endswith("-preview")]
    
    print("稳定版本:")
    for v in sorted(stable_versions, reverse=True):
        print(f"  • {v}" + (" (当前)" if v == current_api_version else ""))
    
    print("\n预览版本:")
    for v in sorted(preview_versions, reverse=True):
        print(f"  • {v}" + (" (当前)" if v == current_api_version else ""))
    
    # 如果需要显示详细信息
    if args.detail:
        print("\n各版本主要特性:")
        for version in API_VERSION_FEATURES:
            display_version_features(version)
    else:
        print("\n提示: 使用 --detail 参数可显示各版本的功能特性")
        print("      使用 --version <版本号> 可查看特定版本的详情")

if __name__ == "__main__":
    main()

#!/usr/bin/env python
# filepath: c:\Users\kevin.li\OneDrive - GREEN DOT CORPORATION\Documents\GitHub\Azure\AzureAI\main_fixed_corrected.py

import os
import json
import argparse
import requests
from dotenv import load_dotenv

# 创建 utilities 文件夹 (如果不存在)
os.makedirs("utilities", exist_ok=True)

class AzureOpenAI:
    """
    Azure OpenAI 服务集成类
    """    
    def __init__(self):
        # 从环境变量加载配置
        # 显示环境变量的值，便于调试
        print("尝试从环境变量读取配置...")
        self.api_key = os.environ.get("AZURE_OPENAI_KEY", "")
        self.endpoint = os.environ.get("AZURE_OPENAI_ENDPOINT", "")
        self.api_version = os.environ.get("AZURE_OPENAI_API_VERSION", "2023-05-15")
        self.deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-35-turbo")  # 部署名称
        
        # 验证配置
        if not self.api_key or not self.endpoint:
            print("警告: 未设置 Azure OpenAI 服务的API密钥或终端点")
            print("请在 .env 文件中设置 AZURE_OPENAI_KEY 和 AZURE_OPENAI_ENDPOINT")
            print("(.env 文件路径: " + os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env') + ")")
    
    def check_config(self):
        """检查配置是否有效"""
        print("\nAzure OpenAI 配置状态:")
        print(f"API终端点: {self.endpoint}")
        print(f"API密钥: {'已设置' if self.api_key else '未设置'}")
        print(f"API版本: {self.api_version}")
        print(f"部署名称: {self.deployment}")
        
        # 检查是否可以连接到Azure
        try:
            if self.api_key and self.endpoint:
                # 首先检查可用的模型
                url = f"{self.endpoint}/openai/models?api-version={self.api_version}"
                headers = {
                    "api-key": self.api_key,
                    "Content-Type": "application/json"
                }
                response = requests.get(url, headers=headers)
                if response.status_code == 200:
                    print("连接测试: 成功 ✓")
                    
                    # 检查部署名称是否有效
                    models = response.json().get("data", [])
                    model_ids = [model.get("id") for model in models]
                    print(f"可用的模型: {', '.join(model_ids) if model_ids else '未发现模型'}")
                    
                    if self.deployment not in model_ids:
                        print(f"警告: 部署名称 '{self.deployment}' 可能不是有效的模型ID")
                        print("您可能需要检查部署名称，或者在Azure门户网站上检查您的部署配置")
                    
                    return True
                else:
                    print(f"连接测试: 失败 ✗ (状态码: {response.status_code})")
                    print(f"错误: {response.text}")
                    return False
            else:
                print("连接测试: 跳过 (缺少API密钥或终端点)")
                return False
        except Exception as e:
            print(f"连接测试: 失败 ✗ (错误: {str(e)})")
            return False
    def chat_completion(self, messages, temperature=0.7, max_tokens=800):
        """
        使用 Azure OpenAI 的 Chat Completion API
        
        参数:
            messages (list): 消息列表，格式为 [{"role": "user", "content": "你好"}]
            temperature (float): 生成文本的随机性程度 (0-1)
            max_tokens (int): 生成的最大令牌数
            
        返回:
            dict: API 响应结果
        """
        try:
            # 确保endpoint没有末尾斜杠
            endpoint = self.endpoint.rstrip('/')
            url = f"{endpoint}/openai/deployments/{self.deployment}/chat/completions?api-version={self.api_version}"
            headers = {
                "api-key": self.api_key,
                "Content-Type": "application/json"
            }
            
            # 移除冗余的model参数，因为已经在URL中指定了deployment
            data = {
                "messages": messages,
                "temperature": temperature,
                "max_tokens": max_tokens
            }
            print(f"发送请求到: {url}")
            print(f"请求数据: {json.dumps(data, ensure_ascii=False)}")
            print(f"使用的API版本: {self.api_version}")
            print(f"使用的部署名称: {self.deployment}")
            
            response = requests.post(url, headers=headers, json=data)
            print(f"响应状态码: {response.status_code}")
            print(f"响应头: {response.headers}")
            
            try:
                result = response.json()
            except json.JSONDecodeError:
                print(f"无法解析JSON响应: {response.text}")
                return {"error": "Invalid JSON response", "response_text": response.text}
            
            if response.status_code == 200:
                if "choices" in result and len(result["choices"]) > 0:
                    content = result["choices"][0]["message"]["content"]
                    print(f"生成的回复: {content[:100]}...")
                return result
            else:
                error_message = "未知错误"
                if isinstance(result, dict) and "error" in result:
                    error_message = result["error"]
                print(f"错误: {response.text}")
                print(f"详细错误信息: {error_message}")
                return {"error": response.text}
        
        except Exception as e:
            print(f"调用 Azure OpenAI 时发生错误: {str(e)}")
            return {"error": str(e)}
    
    def text_completion(self, prompt, temperature=0.7, max_tokens=800):
        """
        使用 Azure OpenAI 的 Completion API
        
        参数:
            prompt (str): 提示文本
            temperature (float): 生成文本的随机性程度 (0-1)
            max_tokens (int): 生成的最大令牌数
            
        返回:
            dict: API 响应结果
        """
        try:
            url = f"{self.endpoint}/openai/deployments/{self.deployment}/completions?api-version={self.api_version}"
            headers = {
                "api-key": self.api_key,
                "Content-Type": "application/json"
            }
            
            data = {
                "prompt": prompt,
                "temperature": temperature,
                "max_tokens": max_tokens,
                "model": self.deployment  # 添加模型参数
            }
            
            print(f"发送请求到: {url}")
            print(f"提示: {prompt[:50]}...")
            
            response = requests.post(url, headers=headers, json=data)
            result = response.json()
            
            if response.status_code == 200:
                if "choices" in result and len(result["choices"]) > 0:
                    text = result["choices"][0]["text"]
                    print(f"生成的文本: {text[:100]}...")
                return result
            else:
                print(f"错误: {response.text}")
                return {"error": response.text}
        
        except Exception as e:
            print(f"调用 Azure OpenAI 时发生错误: {str(e)}")
            return {"error": str(e)}
    
    def embeddings(self, text):
        """
        使用 Azure OpenAI 的 Embeddings API 获取文本的向量表示
        
        参数:
            text (str): 输入文本
            
        返回:
            dict: 包含文本嵌入向量的结果
        """
        try:
            url = f"{self.endpoint}/openai/deployments/{self.deployment}/embeddings?api-version={self.api_version}"
            headers = {
                "api-key": self.api_key,
                "Content-Type": "application/json"
            }
            
            data = {
                "input": text,
                "model": self.deployment  # 添加模型参数
            }
            
            print(f"获取以下文本的嵌入向量: {text[:50]}...")
            
            response = requests.post(url, headers=headers, json=data)
            result = response.json()
            
            if response.status_code == 200:
                print("成功获取嵌入向量")
                return result
            else:
                print(f"错误: {response.text}")
                return {"error": response.text}
        
        except Exception as e:
            print(f"获取嵌入向量时发生错误: {str(e)}")
            return {"error": str(e)}      
    def verify_deployment(self):
        """
        验证部署名称是否有效
        
        返回:
            bool: 部署名称是否有效
        """
        try:
            # 获取可用的部署列表
            endpoint = self.endpoint.rstrip('/')
            url = f"{endpoint}/openai/deployments?api-version={self.api_version}"
            headers = {
                "api-key": self.api_key,
                "Content-Type": "application/json"
            }
            
            print(f"验证部署名称 '{self.deployment}'...")
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                result = response.json()
                deployments = []
                
                # 解析部署列表
                if "data" in result:
                    deployments = [item.get("id") for item in result.get("data", [])]
                
                print(f"可用的部署: {', '.join(deployments) if deployments else '未找到部署'}")
                
                if self.deployment in deployments:
                    print(f"部署验证成功: '{self.deployment}' 是有效的部署")
                    return True
                else:
                    print(f"部署验证失败: '{self.deployment}' 不是有效的部署")
                    return False
            else:
                print(f"无法获取部署列表 (状态码: {response.status_code})")
                print(f"错误: {response.text}")
                return False
                
        except Exception as e:
            print(f"验证部署时发生错误: {str(e)}")
            return False
            
    def list_api_versions(self):
        """
        获取Azure OpenAI支持的API版本信息
        
        返回:
            list: 支持的API版本列表
        """
        try:
            # 这个端点可能会根据Azure OpenAI服务的更新而变化
            # 这里使用一个通用的元数据端点来获取信息
            url = f"{self.endpoint}/openai/info?api-version={self.api_version}"
            headers = {
                "api-key": self.api_key,
                "Content-Type": "application/json"
            }
            
            print(f"尝试获取API版本信息...")
            
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                result = response.json()
                print("成功获取API信息")
                
                # 解析版本信息（具体结构可能需要根据实际响应调整）
                if "api_versions" in result:
                    versions = result["api_versions"]
                    print(f"可用的API版本: {', '.join(versions)}")
                    return versions
                else:
                    print("响应中未找到API版本信息")
                    return []
            else:
                # 如果直接查询不支持，提供一个静态的列表作为备选
                print(f"无法直接获取API版本信息 (状态码: {response.status_code})")
                print("提供已知的API版本列表...")
                known_versions = [
                    "2025-05-15", "2024-07-01-preview", "2024-05-01", 
                    "2023-12-01-preview", "2023-09-15-preview", 
                    "2023-07-01-preview", "2023-06-01-preview", 
                    "2023-05-15", "2022-12-01"
                ]
                print(f"已知的API版本: {', '.join(known_versions)}")
                return known_versions
        
        except Exception as e:
            print(f"获取API版本信息时发生错误: {str(e)}")
            return []


def create_dotenv_file():
    """创建.env文件模板（如果不存在）"""
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
    
    if not os.path.exists(env_path):
        template = """# Azure OpenAI 配置
AZURE_OPENAI_KEY=你的API密钥
AZURE_OPENAI_ENDPOINT=https://你的资源名称.openai.azure.com/
AZURE_OPENAI_API_VERSION=2023-05-15
AZURE_OPENAI_DEPLOYMENT=你的部署名称
"""        
        try:
            with open(env_path, 'w', encoding='utf-8') as f:
                f.write(template)
            print(f"已创建 .env 模板文件: {env_path}")
            print("请编辑该文件，填写您的 Azure OpenAI 服务凭据")
            return True
        except Exception as e:
            print(f"创建 .env 文件时出错: {str(e)}")
            return False
    else:
        print(f".env 文件已存在: {env_path}")
        return True


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="Azure OpenAI 服务集成示例")    
    parser.add_argument("--mode", choices=["chat", "completion", "embeddings", "test", "versions"], 
                        help="选择操作模式", default="chat")
    parser.add_argument("--input", help="输入文本或提示", default="")
    parser.add_argument("--output", help="输出文件路径")
    parser.add_argument("--temp", type=float, help="温度参数 (0-1)", default=0.7)
        
    args = parser.parse_args()
    
    # 创建.env文件（如果不存在）
    create_dotenv_file()
    
    # 加载环境变量
    load_dotenv()
    
    # 初始化 Azure OpenAI 服务
    azure_openai = AzureOpenAI()
    print("args.mode=", args.mode)
    
    # 如果是测试模式，只检查配置
    if args.mode == "test":
        print("运行配置测试...")
        azure_openai.check_config()
        return
        
    # 如果是版本信息模式，获取API版本信息
    if args.mode == "versions":
        print("获取API版本信息...")
        azure_openai.list_api_versions()
        return
    
    # 获取输入内容
    input_text = args.input
    if not input_text:
        input_text = input("请输入提示或文本: ")
    
    # 处理不同的模式
    result = None
    if args.mode == "chat":
        messages = [{"role": "user", "content": input_text}]
        result = azure_openai.chat_completion(messages, temperature=args.temp)
    elif args.mode == "completion":
        result = azure_openai.text_completion(input_text, temperature=args.temp)
    elif args.mode == "embeddings":
        result = azure_openai.embeddings(input_text)
    
    # 打印结果
    if result:
        print("\n结果:")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        
        # 保存到文件（如果提供了输出路径）
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            print(f"结果已保存到文件: {args.output}")


if __name__ == "__main__":
    main()

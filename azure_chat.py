import os
import json
import argparse
import requests
from dotenv import load_dotenv  
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
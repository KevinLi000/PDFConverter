
import os
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
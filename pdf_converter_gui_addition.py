    def _update_status(self, message, progress):
        """更新状态和进度条"""
        self.status_var.set(message)
        self.progress_var.set(progress)
        
    def _check_conversion_result(self):
        """检查转换结果并处理"""
        try:
            if not self.conversion_queue.empty():
                status, result = self.conversion_queue.get()
                
                if status == "success":
                    # 转换成功
                    # 询问是否打开输出文件
                    if messagebox.askyesno("转换成功", f"PDF已成功转换！是否打开输出文件？\n{result}"):
                        try:
                            os.startfile(result)
                        except Exception as e:
                            messagebox.showerror("错误", f"无法打开文件: {str(e)}")
                else:  # error
                    # 显示错误信息
                    messagebox.showerror("转换失败", f"转换过程中出现错误：\n{result}")
                
                # 重置按钮状态
                self.convert_button.config(state="normal")
            else:
                # 队列为空，继续等待
                self.root.after(100, self._check_conversion_result)
        except Exception as e:
            messagebox.showerror("错误", f"处理转换结果时出错: {str(e)}")
            self.convert_button.config(state="normal")

# 如果直接运行此脚本
if __name__ == "__main__":
    try:
        # 创建主窗口
        root = tk.Tk()
        app = PDFConverterGUI(root)
        
        # 开始主事件循环
        root.mainloop()
    except Exception as e:
        messagebox.showerror("错误", f"程序启动失败: {str(e)}")

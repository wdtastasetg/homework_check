import sys
import os
import subprocess

def run_script(script_name):
    """运行指定的Python脚本"""
    script_path = os.path.join(os.getcwd(), script_name)
    if not os.path.exists(script_path):
        print(f"错误: 找不到脚本 {script_name}")
        return
    
    print(f"\n>>> 正在启动 {script_name} ...")
    try:
        # 使用当前 Python 解释器运行脚本
        subprocess.run([sys.executable, script_name], check=False)
    except Exception as e:
        print(f"运行脚本出错: {e}")
    print(f"<<< {script_name} 运行结束\n")

def main():
    while True:
        print("=" * 40)
        print("      作业管理系统主菜单")
        print("=" * 40)
        print("1. 提取名单 (extract.py) - 从Excel提取学生名单")
        print("2. 下载作业 (mail.py)    - 从邮箱下载作业附件")
        print("3. 检查提交 (check.py)   - 检查作业提交情况")
        print("4. 批改作业 (grade.py)   - 逐个打开文件进行评分")
        print("0. 退出程序")
        print("-" * 40)
        
        choice = input("请输入功能序号 (0-4): ").strip()
        
        if choice == '1':
            run_script("extract.py")
        elif choice == '2':
            run_script("mail.py")
        elif choice == '3':
            run_script("check.py")
        elif choice == '4':
            run_script("grade.py")
        elif choice == '0':
            print("感谢使用，再见！")
            break
        else:
            print("无效输入，请重新选择。")

if __name__ == "__main__":
    main()

import os
import json
import subprocess
import sys
import platform

def open_file(filepath):
    """
    根据操作系统打开文件
    """
    if platform.system() == 'Darwin':       # macOS
        subprocess.call(('open', filepath))
    elif platform.system() == 'Windows':    # Windows
        os.startfile(filepath)
    else:                                   # linux variants
        subprocess.call(('xdg-open', filepath))

def save_grades(grades, output_file):
    """
    保存成绩到JSON文件
    """
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(grades, f, ensure_ascii=False, indent=2)
        print(f"成绩已保存到: {output_file}")
    except Exception as e:
        print(f"保存文件失败: {e}")

def main():
    print("作业批改助手")
    print("=" * 50)
    
    # 1. 输入文件夹名称
    folder_name = input("请输入要批改的作业文件夹名称 (位于 downloads 目录下): ").strip()
    if not folder_name:
        print("未输入文件夹名称，程序退出")
        return
    
    folder_path = os.path.join("downloads", folder_name)
    
    if not os.path.exists(folder_path):
        print(f"错误：文件夹 '{folder_path}' 不存在。")
        return
    
    # 2. 获取文件列表
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    # 过滤掉隐藏文件或非作业文件（可选）
    files = [f for f in files if not f.startswith('.')]
    files.sort() # 排序，保证顺序一致
    
    if not files:
        print("文件夹为空，没有可批改的文件。")
        return
    
    print(f"共找到 {len(files)} 个文件，开始批改...")
    print("-" * 50)
    
    grades = []
    
    # 3. 循环批改
    for i, filename in enumerate(files):
        filepath = os.path.join(folder_path, filename)
        print(f"[{i+1}/{len(files)}] 正在打开: {filename}")
        
        # 自动打开文件
        try:
            open_file(filepath)
        except Exception as e:
            print(f"无法自动打开文件: {e}")
        
        # 输入分数
        while True:
            score_input = input(f"请输入 '{filename}' 的分数 (输入 'q' 退出, 's' 跳过): ").strip()
            
            if score_input.lower() == 'q':
                print("批改中断，正在保存已批改内容...")
                output_file = f"{folder_name}_grades.json"
                save_grades(grades, output_file)
                return
            
            if score_input.lower() == 's':
                print("已跳过")
                break
                
            try:
                score = score_input
                
                # 记录结果
                grades.append({
                    "filename": filename,
                    "score": score
                })
                break
            except Exception:
                print("输入无效，请重新输入")
        
        print("-" * 30)

    # 4. 保存结果
    output_file = f"{folder_name}_grades.json"
    save_grades(grades, output_file)
    print("所有文件批改完成！")

if __name__ == "__main__":
    main()

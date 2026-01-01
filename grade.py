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
        # 优先尝试使用 VS Code 的 code 命令打开 (适配远程开发环境)
        try:
            # 使用 code 命令打开文件
            subprocess.call(('code', filepath))
        except FileNotFoundError:
            # 如果没有 code 命令，则尝试 xdg-open
            try:
                subprocess.call(('xdg-open', filepath))
            except FileNotFoundError:
                print(f"无法打开文件: {filepath} (未找到打开方式)")

def save_grades(grades, output_file):
    """
    保存成绩到JSON文件
    """
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(grades, f, ensure_ascii=False, indent=2)
        print(f"成绩已保存到: {output_file}")
        
        # 计算并显示平均分
        total_score = 0
        valid_count = 0
        for record in grades:
            try:
                score_val = float(record.get('score', 0))
                total_score += score_val
                valid_count += 1
            except (ValueError, TypeError):
                pass
                
        if valid_count > 0:
            average = total_score / valid_count
            print(f"当前平均分: {average:.2f} (基于 {valid_count} 份有效成绩)")
            
    except Exception as e:
        print(f"保存文件失败: {e}")

def load_students():
    """
    加载学生名单
    """
    try:
        if os.path.exists('students.json'):
            with open('students.json', 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return []

def find_student_in_list(filename, students):
    """
    在学生名单中查找匹配的学生
    """
    for student in students:
        student_id = str(student.get('id', ''))
        student_name = str(student.get('name', ''))
        
        # 优先匹配学号 (如果学号存在且在文件名中)
        if student_id and student_id in filename:
            return student
        
        # 其次匹配姓名 (如果姓名存在且在文件名中，且长度大于1以避免误判)
        if student_name and len(student_name) > 1 and student_name in filename:
            return student
            
    return None

def main():
    print("作业批改助手")
    print("=" * 50)
    
    # 加载学生名单
    students = load_students()
    if students:
        print(f"已加载 {len(students)} 名学生信息用于匹配。")
    else:
        print("未找到 students.json，将只记录文件名和分数。")

    # 1. 获取downloads下的所有文件夹
    downloads_path = "downloads"
    if not os.path.exists(downloads_path):
        print(f"错误：目录 '{downloads_path}' 不存在。")
        return

    subfolders = [f for f in os.listdir(downloads_path) if os.path.isdir(os.path.join(downloads_path, f))]
    subfolders.sort()

    if not subfolders:
        print(f"'{downloads_path}' 下没有找到任何文件夹。")
        return

    print("请选择要批改的作业文件夹:")
    for i, folder in enumerate(subfolders, 1):
        print(f"{i}. {folder}")
    
    choice = input("请输入序号: ").strip()
    
    if not choice.isdigit():
        print("输入无效，请输入数字。")
        return
    
    idx = int(choice) - 1
    if idx < 0 or idx >= len(subfolders):
        print("输入的序号超出范围。")
        return
        
    folder_name = subfolders[idx]
    print(f"已选择: {folder_name}")
    
    folder_path = os.path.join(downloads_path, folder_name)
    
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
    
    # 尝试加载已有成绩
    output_file = f"{folder_name}_grades.json"
    grades = []
    graded_filenames = set()
    
    if os.path.exists(output_file):
        try:
            with open(output_file, 'r', encoding='utf-8') as f:
                grades = json.load(f)
                for record in grades:
                    if 'filename' in record:
                        graded_filenames.add(record['filename'])
            print(f"已加载 {len(grades)} 条历史成绩，将跳过已评分文件。")
        except Exception as e:
            print(f"加载历史成绩失败: {e}")
    
    word_hint_shown = False
    
    # 3. 循环批改
    for i, filename in enumerate(files):
        if filename in graded_filenames:
            print(f"[{i+1}/{len(files)}] {filename} (已评分) - 跳过")
            continue

        filepath = os.path.join(folder_path, filename)
        print(f"[{i+1}/{len(files)}] 正在打开: {filename}")
        
        # 检查是否为 Word 文档并提示
        if filename.lower().endswith(('.doc', '.docx')) and not word_hint_shown:
            print("提示: VS Code 默认无法预览 Word 文档。")
            print("      请在扩展商店搜索并安装 'Office Viewer' (cweijan.vscode-office) 插件以正常查看。")
            word_hint_shown = True

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
                
                # 尝试匹配学生信息
                matched_student = find_student_in_list(filename, students)
                
                if matched_student:
                    # 如果匹配成功，基于学生信息构建记录
                    record = matched_student.copy()
                    record['score'] = score
                    record['filename'] = filename # 保留文件名作为参考
                    grades.append(record)
                    print(f"  -> 已匹配学生: {matched_student.get('name')} ({matched_student.get('id')})")
                else:
                    # 无法匹配，按原格式记录
                    grades.append({
                        "filename": filename,
                        "score": score
                    })
                    print("  -> 未匹配到学生信息，仅记录文件名。")
                
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

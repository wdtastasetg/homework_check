import os
import json
from typing import List, Tuple

def load_student_list(json_file: str) -> List[dict]:
    """从JSON文件加载学生名单"""
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            students = json.load(f)
        print(f"成功加载 {len(students)} 名学生信息")
        return students
    except Exception as e:
        print(f"加载JSON文件失败: {e}")
        return []

def get_files_in_folder(folder_path: str) -> List[str]:
    """获取文件夹中的所有文件名"""
    if not os.path.exists(folder_path):
        print(f"文件夹不存在: {folder_path}")
        return []
    
    files = []
    for filename in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, filename)):
            files.append(filename)
    
    print(f"在文件夹中找到 {len(files)} 个文件")
    return files

def check_submission(student_id: str, student_name: str, files: List[str]) -> Tuple[bool, str]:
    """检查单个学生的作业提交情况"""
    for filename in files:
        # 检查文件名是否包含学号和姓名
        if student_id in filename and student_name in filename:
            return True, filename
        # 宽松匹配：只检查学号
        elif student_id in filename:
            return True, filename
        # 更宽松的匹配：检查姓名（可能存在重名风险）
        elif student_name in filename and len(student_name) > 1:
            return True, filename
    return False, ""

def analyze_all_submissions(students: List[dict], folder_path: str) -> Tuple[List[dict], List[dict]]:
    """分析所有学生的提交情况"""
    files = get_files_in_folder(folder_path)
    if not files:
        return [], students
    
    submitted = []
    not_submitted = []
    
    for student in students:
        student_id = student['id']
        student_name = student['name']
        
        has_submitted, filename = check_submission(student_id, student_name, files)
        
        if has_submitted:
            submitted.append({
                'serial': student['serial'],
                'id': student_id,
                'name': student_name,
                'filename': filename
            })
        else:
            not_submitted.append({
                'serial': student['serial'],
                'id': student_id,
                'name': student_name
            })
    
    return submitted, not_submitted

def generate_report(submitted: List[dict], not_submitted: List[dict], output_file: str = None) -> str:
    """生成提交情况报告"""
    total_students = len(submitted) + len(not_submitted)
    submission_rate = (len(submitted) / total_students) * 100 if total_students > 0 else 0
    
    report = []
    report.append("=" * 60)
    report.append("作业提交情况统计报告")
    report.append("=" * 60)
    report.append(f"总学生数: {total_students}")
    report.append(f"已提交: {len(submitted)} 人")
    report.append(f"未提交: {len(not_submitted)} 人")
    report.append(f"提交率: {submission_rate:.1f}%")
    report.append("")
    
    # 已提交学生名单
    report.append("已提交学生名单:")
    report.append("-" * 40)
    for student in sorted(submitted, key=lambda x: x['serial']):
        report.append(f"{student['serial']:3d}. {student['id']} {student['name']:10s} - {student['filename']}")
    
    report.append("")
    
    # 未提交学生名单
    report.append("未提交学生名单:")
    report.append("-" * 40)
    for student in sorted(not_submitted, key=lambda x: x['serial']):
        report.append(f"{student['serial']:3d}. {student['id']} {student['name']}")
    
    report_text = "\n".join(report)
    
    # 输出到文件
    if output_file:
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(report_text)
            print(f"报告已保存到: {output_file}")
        except Exception as e:
            print(f"保存报告文件失败: {e}")
    
    return report_text

def main():
    """主函数"""
    print("作业提交检查程序")
    print("=" * 50)
    
    # 加载学生名单
    students = load_student_list("students.json")
    if not students:
        return
    
    # 输入作业文件夹名称
    folder_name = input("请输入作业文件夹名称: ").strip()
    if not folder_name:
        print("未输入文件夹名称，程序退出")
        return
    
    # 构造完整路径
    folder_path = os.path.join("downloads", folder_name)

    # 分析提交情况
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print(f"错误：文件夹 '{folder_path}' 不存在。")
        return

    submitted, not_submitted = analyze_all_submissions(students, folder_path)
    
    # 设置默认输出文件名
    default_output_file = f"{folder_name}_output.txt"
    
    # 询问用户是否使用默认输出文件名
    use_default = input(f"是否使用默认输出文件名 '{default_output_file}'? (y/n, 回车默认使用): ").strip().lower()
    
    if use_default == 'n':
        output_file = input("请输入自定义报告输出文件名: ").strip()
        if not output_file:
            output_file = default_output_file
    else:
        output_file = default_output_file
    
    # 生成报告
    report = generate_report(submitted, not_submitted, output_file)

if __name__ == "__main__":
    main()
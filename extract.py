import pandas as pd
import json
import re

def extract_students_from_excel(excel_file):
    """
    从Excel文件中提取学生信息
    
    Args:
        excel_file (str): Excel文件路径
        
    Returns:
        list: 学生信息列表
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file, sheet_name='RollBook (2)', header=None)
        
        students = []
        
        # 遍历所有行
        for index, row in df.iterrows():
            # 检查是否为学生数据行（序号列为数字）
            if pd.notna(row[0]) and str(row[0]).isdigit():
                serial = int(row[0])
                student_id = str(row[1]).strip() if pd.notna(row[1]) else ""
                name = str(row[2]).strip() if pd.notna(row[2]) else ""
                
                # 清理姓名中的特殊标记（如*号）
                name = re.sub(r'\s*\*', '', name)
                
                # 确保所有必要字段都存在
                if student_id and name:
                    student = {
                        "serial": serial,
                        "id": student_id,
                        "name": name
                    }
                    students.append(student)
        
        # 按序号排序
        students.sort(key=lambda x: x["serial"])
        
        return students
        
    except Exception as e:
        print(f"读取Excel文件时出错: {e}")
        return []

def save_to_json(students, json_file):
    """
    将学生信息保存为JSON文件
    
    Args:
        students (list): 学生信息列表
        json_file (str): JSON文件路径
    """
    try:
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(students, f, ensure_ascii=False, indent=2)
        print(f"数据已成功保存到 {json_file}")
    except Exception as e:
        print(f"保存JSON文件时出错: {e}")

def main():
    # 文件路径
    default_file = "总名单.xlsx"
    # 获取用户输入的文件名，如果为空则使用默认值
    user_input = input(f"请输入Excel文件名 (默认为 {default_file}): ").strip()
    excel_file = user_input if user_input else default_file
    
    output_json = "students.json"
    
    # 从Excel提取数据
    print("正在从Excel文件中提取学生信息...")
    students = extract_students_from_excel(excel_file)
    
    if students:
        print(f"成功提取 {len(students)} 名学生信息")
        
        # 显示前几条记录作为示例
        print("\n前5条记录示例:")
        for i, student in enumerate(students[:5]):
            print(f"序号: {student['serial']}, 学号: {student['id']}, 姓名: {student['name']}")
        
        # 保存为JSON文件
        save_to_json(students, output_json)
        
        # 同时在控制台输出JSON内容
        print(f"\n生成的JSON内容:")
        print(json.dumps(students, ensure_ascii=False, indent=2))
    else:
        print("未找到学生数据")

if __name__ == "__main__":
    main()
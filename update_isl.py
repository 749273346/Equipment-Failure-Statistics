
import os

additional_content = """
WizardSelectTasks=选择附加任务
SelectTasksDesc=您想要安装程序执行哪些附加任务？
SelectTasksLabel=选择您想要安装程序在安装 [name] 时执行的附加任务，然后点击“下一步”。
WizardInstalling=正在安装
InstallingLabel=正在安装 [name]，请稍候。
"""

file_path = r"E:\QC-攻关小组\正在进行项目\设备故障统计\ChineseSimplified.isl"

# Read existing content
with open(file_path, 'r', encoding='utf-8-sig') as f:
    content = f.read()

# Append if not present (simple check)
if "WizardSelectTasks=" not in content:
    with open(file_path, 'w', encoding='utf-8-sig') as f:
        f.write(content + additional_content)
    print("Updated ChineseSimplified.isl")
else:
    print("Content already present")

import os
import win32com.client
import gc
import sys
from pathlib import Path

# 解决中文路径问题
import pythoncom

pythoncom.CoInitialize()


def validate_path(file_path):
    """验证路径是否存在"""
    if not os.path.exists(file_path):
        print(f"\n【错误】路径不存在：{file_path}")
        return False
    if not os.access(file_path, os.W_OK):
        print(f"\n【错误】路径没有写入权限：{file_path}")
        return False
    return True


def change_suffix_to_pdf(file_name):
    """修改文件后缀为PDF"""
    if '.' not in file_name:
        return file_name + ".pdf"
    return file_name[:file_name.rfind('.')] + ".pdf"


def add_worksheet_suffix(file_name, sheet_num):
    """为Excel工作表添加序号后缀"""
    if '.' not in file_name:
        return f"{file_name}_工作表{sheet_num}.pdf"
    return file_name[:file_name.rfind('.')] + f"_工作表{sheet_num}.pdf"


def get_output_path(file_path, file_name):
    """获取PDF输出路径"""
    pdf_folder = os.path.join(file_path, 'pdf')
    Path(pdf_folder).mkdir(exist_ok=True)  # 安全创建文件夹
    pdf_name = change_suffix_to_pdf(file_name)
    return os.path.join(pdf_folder, pdf_name)


def word_to_pdf(file_path, word_files):
    """Word转PDF"""
    if not word_files:
        print("\n【无 Word 文件】")
        return

    print("\n【开始 Word -> PDF 转换】")
    word_app = None
    doc = None

    try:
        # 优先使用Microsoft Word，备用WPS
        try:
            word_app = win32com.client.Dispatch("Word.Application")
        except:
            word_app = win32com.client.Dispatch("Kwps.Application")

        word_app.Visible = 0
        word_app.DisplayAlerts = 0

        for idx, file_name in enumerate(word_files):
            print(f"\n[{idx + 1}/{len(word_files)}] 处理：{file_name}")
            full_input_path = os.path.join(file_path, file_name)
            full_output_path = get_output_path(file_path, file_name)

            try:
                doc = word_app.Documents.Open(full_input_path)
                doc.SaveAs(full_output_path, FileFormat=17)  # 17 = PDF格式
                print(f"✅ 转换完成：{os.path.basename(full_output_path)}")
            except Exception as e:
                print(f"❌ 转换失败：{file_name} - {str(e)}")
            finally:
                if doc:
                    doc.Close(SaveChanges=0)
                    doc = None

    except Exception as e:
        print(f"\n【Word转换异常】：{str(e)}")
    finally:
        if word_app:
            word_app.Quit()
        gc.collect()
        print("\n【Word 转换进程已结束】")


def excel_to_pdf(file_path, excel_files):
    """Excel转PDF（按工作表拆分）"""
    if not excel_files:
        print("\n【无 Excel 文件】")
        return

    print("\n【开始 Excel -> PDF 转换】")
    excel_app = None
    wb = None
    ws = None

    try:
        # 优先使用Microsoft Excel，备用WPS
        try:
            excel_app = win32com.client.Dispatch("Excel.Application")
        except:
            excel_app = win32com.client.Dispatch("ET.Application")

        excel_app.Visible = 0
        excel_app.DisplayAlerts = 0

        for idx, file_name in enumerate(excel_files):
            print(f"\n[{idx + 1}/{len(excel_files)}] 处理：{file_name}")
            full_input_path = os.path.join(file_path, file_name)

            try:
                wb = excel_app.Workbooks.Open(full_input_path)
                total_sheets = wb.Worksheets.Count

                for sheet_idx in range(total_sheets):
                    sheet_num = sheet_idx + 1
                    ws = wb.Worksheets(sheet_num)

                    # 生成带工作表序号的文件名
                    base_name = file_name[:file_name.rfind('.')] if '.' in file_name else file_name
                    pdf_name = f"{base_name}_工作表{sheet_num}.pdf"
                    full_output_path = os.path.join(file_path, 'pdf', pdf_name)

                    ws.ExportAsFixedFormat(0, full_output_path)  # 0 = PDF格式
                    print(f"✅ 工作表{sheet_num}转换完成：{pdf_name}")

            except Exception as e:
                print(f"❌ 转换失败：{file_name} - {str(e)}")
            finally:
                if wb:
                    wb.Close(SaveChanges=0)
                    wb = None
                ws = None

    except Exception as e:
        print(f"\n【Excel转换异常】：{str(e)}")
    finally:
        if excel_app:
            excel_app.Quit()
        gc.collect()
        print("\n【Excel 转换进程已结束】")


def ppt_to_pdf(file_path, ppt_files):
    """PPT转PDF"""
    if not ppt_files:
        print("\n【无 PPT 文件】")
        return

    print("\n【开始 PPT -> PDF 转换】")
    ppt_app = None
    presentation = None

    try:
        # 优先使用Microsoft PowerPoint，备用WPS
        try:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        except:
            ppt_app = win32com.client.Dispatch("WPP.Application")

        ppt_app.Visible = 0

        for idx, file_name in enumerate(ppt_files):
            print(f"\n[{idx + 1}/{len(ppt_files)}] 处理：{file_name}")
            full_input_path = os.path.join(file_path, file_name)
            full_output_path = get_output_path(file_path, file_name)

            try:
                # WithWindow=False 避免显示PPT窗口
                presentation = ppt_app.Presentations.Open(
                    full_input_path,
                    WithWindow=False,
                    ReadOnly=True
                )

                if presentation.Slides.Count == 0:
                    print(f"⚠️ 跳过空文件：{file_name}")
                    continue

                presentation.SaveAs(full_output_path, 32)  # 32 = PDF格式
                print(f"✅ 转换完成：{os.path.basename(full_output_path)}")

            except Exception as e:
                print(f"❌ 转换失败：{file_name} - {str(e)}")
            finally:
                if presentation:
                    presentation.Close()
                    presentation = None

    except Exception as e:
        print(f"\n【PPT转换异常】：{str(e)}")
    finally:
        if ppt_app:
            ppt_app.Quit()
        gc.collect()
        print("\n【PPT 转换进程已结束】")


def collect_files(file_path):
    """收集指定路径下的Office文件"""
    files = {
        'word': [],
        'excel': [],
        'ppt': []
    }

    for file_name in os.listdir(file_path):
        # 跳过隐藏文件和文件夹
        if file_name.startswith('.') or os.path.isdir(os.path.join(file_path, file_name)):
            continue

        # 按后缀分类
        lower_name = file_name.lower()
        if lower_name.endswith(('.doc', '.docx')):
            files['word'].append(file_name)
        elif lower_name.endswith(('.xls', '.xlsx')):
            files['excel'].append(file_name)
        elif lower_name.endswith(('.ppt', '.pptx')):
            files['ppt'].append(file_name)

    return files


def main():
    """主程序"""
    print("=" * 50)
    print("Office 文件转 PDF 工具")
    print("=" * 50)
    print("功能：将指定路径下的Word/Excel/PPT转换为PDF")
    print("支持格式：")
    print("  Word: .doc, .docx")
    print("  Excel: .xls, .xlsx (按工作表拆分)")
    print("  PPT: .ppt, .pptx")
    print("=" * 50)

    # 获取并验证用户输入的路径
    default_path = os.getcwd()
    input_path = input(f"\n请输入目标路径（直接回车使用当前路径：{default_path}）：\n").strip()
    file_path = input_path if input_path else default_path

    # 路径验证
    if not validate_path(file_path):
        input("\n按回车键退出...")
        return

    # 收集文件
    print(f"\n正在扫描 {file_path} 中的Office文件...")
    files = collect_files(file_path)

    # 显示统计信息
    total_files = len(files['word']) + len(files['excel']) + len(files['ppt'])
    print(f"\n文件扫描完成：")
    print(f"  Word 文件：{len(files['word'])} 个")
    print(f"  Excel 文件：{len(files['excel'])} 个")
    print(f"  PPT 文件：{len(files['ppt'])} 个")
    print(f"  总计：{total_files} 个文件")

    if total_files == 0:
        print("\n【未找到任何可转换的Office文件】")
        input("\n按回车键退出...")
        return

    # 确认转换
    confirm = input("\n是否开始转换？(Y/N，默认Y)：").strip().upper()
    if confirm not in ('', 'Y', 'YES'):
        print("\n转换已取消")
        input("\n按回车键退出...")
        return

    # 创建PDF输出文件夹
    pdf_folder = os.path.join(file_path, 'pdf')
    Path(pdf_folder).mkdir(exist_ok=True)
    print(f"\nPDF 文件将保存至：{pdf_folder}")

    # 开始转换
    print("\n" + "=" * 50)
    word_to_pdf(file_path, files['word'])
    excel_to_pdf(file_path, files['excel'])
    ppt_to_pdf(file_path, files['ppt'])
    print("=" * 50)

    print("\n🎉 所有转换任务已完成！")
    print(f"\nPDF 文件保存位置：{pdf_folder}")
    input("\n按回车键退出...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
    except Exception as e:
        print(f"\n\n程序异常：{str(e)}")
        input("\n按回车键退出...")
    finally:
        pythoncom.CoUninitialize()